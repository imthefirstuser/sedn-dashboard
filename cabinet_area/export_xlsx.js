/* export_xlsx.js —— 页面数据 → Excel（纯前端，可在静态页面/Flame 上直接用）
   依赖全局：S / ST / V / pv / parts（index.html 里定义）、XLSXLite（xlsx_lite.js）
   两个导出：exportLive() 公式版（改参数/柜宽自动重算）  exportExcel() 死值快照            */
(function () {
  // ---------- 参数表（Excel「参数」sheet 顺序，name / 页面取值 / 单位 / 说明）----------
  const PARAMS = [
    ["侧板厚", () => pv("side").t, "mm", "柜体板"],
    ["顶板厚", () => pv("top").t, "mm", ""],
    ["底板厚", () => pv("bottom").t, "mm", ""],
    ["层板厚", () => pv("shelf").t, "mm", ""],
    ["竖板厚", () => pv("vdiv").t, "mm", "中隔竖板"],
    ["背板厚", () => pv("back").t, "mm", "薄板"],
    ["见光板厚", () => pv("vis").t, "mm", "外露装饰面"],
    ["抽侧帮厚", () => pv("dhelp").t, "mm", "抽屉"],
    ["抽底板厚", () => pv("dbox").t, "mm", ""],
    ["抽后板厚", () => pv("dback").t, "mm", ""],
    ["门板-平板厚", () => pv("door_p").t, "mm", "平板门"],
    ["门板-造型厚", () => pv("door_m").t, "mm", "造型门"],
    ["门板-玻璃厚", () => pv("door_g").t, "mm", "玻璃门"],
    ["踢脚高", () => V("K"), "mm", "吊柜无踢脚"],
    ["门缝", () => V("G"), "mm", "四周留缝"],
    ["抽屉单侧抽隙", () => V("dGap"), "mm", "三节轨常规"],
    ["抽屉轨道预留", () => V("dRail"), "mm", "抽深=柜深-背板-此项"],
    ["抽体高扣减", () => V("dCut"), "mm", "抽体高=面板高-此项"],
    ["拉篮柜加层板", () => V("bsBoard"), "0/1", "0=拉篮柜不加板（默认） 1=每个拉篮格加一块层板"],
    ["抽屉最小面板高", () => V("dMinFace"), "mm", "约束：抽屉数上限=按此值算（不影响算料）"],
    ["层板最小间距", () => V("spMinGap"), "mm", "约束：层板数上限=按此值算（不影响算料）"],
    ["最小格宽", () => V("bayMin"), "mm", "约束：竖板数上限=按此值算（不影响算料）"],
    ["转角背板片数", () => V("cnBack"), "片", "2=每段各一片(默认)"],
    ["转角背板重叠扣减", () => V("cnOver"), "mm", "两片在转角处扣一个板厚"],
    ["转角W2口径", () => V("cnW2Ref"), "0/1", "0=W2从转角方格外端量(净长) 1=从相邻墙内角量(含方格)"],
    ["门板最低面积-平板", () => V("dmP"), "㎡", "每柜门板合计不足此值按此值计"],
    ["门板最低面积-造型", () => V("dmM"), "㎡", "同上"],
    ["门板最低面积-玻璃", () => V("dmG"), "㎡", "同上"],
    ["门铰·每扇单开", () => V("hjSingle"), "件", "五金：按门扇数"],
    ["门铰·每扇翻门", () => V("hjFlip"), "件", "五金：翻门多用支撑"],
    ["门铰·加铰门高门槛", () => V("hjThresh"), "mm", "五金：门高≤此值只用基础铰数"],
    ["门铰·每加一铰门高步长", () => V("hjStep"), "mm", "五金：每超门槛这么多加 1 个"],
    ["门铰·每扇最多", () => V("hjMax"), "件", "五金：单扇封顶"],
    ["轨道·每个抽屉", () => V("rlDrawer"), "副", "五金：抽屉一副"],
    ["轨道·每个拉篮", () => V("rlBasket"), "副", "五金：拉篮一副"],
    ["门铰单价", () => V("hjPrice"), "元/件", "五金金额"],
    ["轨道单价", () => V("rlPrice"), "元/副", "五金金额"],
  ];
  const UNIT_PRICE = ["侧板", "顶板", "底板", "层板", "竖板", "背板", "门板-平板", "门板-造型", "门板-玻璃",
    "见光板", "见光板-吊柜底", "接顶板", "踢脚板", "转角内侧板", "抽侧帮", "抽底板", "抽后板"];
  const PRICE_ID = { "侧板": "side", "顶板": "top", "底板": "bottom", "层板": "shelf", "竖板": "vdiv", "背板": "back",
    "门板-平板": "door_p", "门板-造型": "door_m", "门板-玻璃": "door_g", "见光板": "vis", "见光板-吊柜底": "vis", "接顶板": "vis", "踢脚板": "kick", "转角内侧板": "corner",
    "抽侧帮": "dhelp", "抽底板": "dbox", "抽后板": "dback" };

  const f = (tpl, min) => `MAX(${tpl},${min == null ? 0 : min})`;   // 尺寸夹到 >=0
  const dnKey = s2 => String(s2).replace(/[^0-9A-Za-z_\u4e00-\u9fa5]/g, "");   // 定义名去 -/·/空格等
  // ---------- 明细模板（与 make_excel.py 的 BOARDS 同口径）----------
  const BAYW = "MAX(G{r}-2*侧板厚-R{r}*竖板厚,0)/(R{r}+1)";   // 普通柜每格宽
  const BOARDS = [
    ["侧板", "侧板厚", "I{r}", 'IF(OR(D{r}="空位",AND(D{r}="转角",X{r}=1)),H{r},H{r}-IF(AND(Y{r}=0,C{r}<>"吊柜"),0,顶板厚)-底板厚)', "2", "侧板", "柜深×高×2：普通柜/直柜 高=柜高-顶-底（顶底板通长），L 柜通高"],
    ["顶板", "顶板厚", "G{r}", 'IF(AND(D{r}="转角",X{r}=1),L{r},I{r})', 'IF(OR(D{r}="空位",AND(Y{r}=0,C{r}<>"吊柜")),0,1)', "顶板", "普通柜/直柜=柜宽×柜深；L 柜=矩形整块 W1×W2（W2 为矩形深度，掏口当废料）"],
    ["底板", "底板厚", "G{r}", 'IF(AND(D{r}="转角",X{r}=1),L{r},I{r})', 'IF(D{r}="空位",0,1)', "底板", "同上"],
    ["背板", "背板厚", 'IF(AND(D{r}="转角",X{r}=1),G{r},G{r}-2*侧板厚)', 'IF(AND(D{r}="转角",X{r}=1),H{r},H{r}-IF(AND(Y{r}=0,C{r}<>"吊柜"),0,顶板厚)-底板厚)', "1", "背板", "L 柜第一片=W1×柜高（通高）；普通柜=内宽×内高"],
    ["背板", "背板厚", 'IF(AND(D{r}="转角",X{r}=1),L{r},0)', 'IF(AND(D{r}="转角",X{r}=1),H{r},H{r}-IF(AND(Y{r}=0,C{r}<>"吊柜"),0,顶板厚)-底板厚)', 'IF(AND(D{r}="转角",X{r}=1,转角背板片数=2),1,0)', "背板", "L 柜第二片=W2×柜高（沿另一面，通高）"],
    ["竖板", "竖板厚", 'I{r}-背板厚', 'H{r}-IF(AND(Y{r}=0,C{r}<>"吊柜"),0,顶板厚)-底板厚', 'IF(OR(D{r}="空位",AND(D{r}="转角",X{r}=1)),0,R{r})', "竖板", "(柜深-背板厚)×内高×竖板数（清单R列）"],
    ["层板", "层板厚", 'IF(AND(D{r}="转角",X{r}=1),G{r},MAX(G{r}-2*侧板厚-R{r}*竖板厚,0)/(R{r}+1))', 'IF(AND(D{r}="转角",X{r}=1),L{r},I{r}-背板厚)', 'IF(AND(D{r}="转角",X{r}=1),J{r},IF(OR(D{r}="层板",AND(D{r}="转角",X{r}=0)),J{r},IF(D{r}="拉篮",拉篮柜加层板,0))*(R{r}+1))', "层板", "格宽×(柜深-背板厚)×层板数×格数；L 柜每层整块 W1×W2（掏口废料）"],
    ["踢脚板", "侧板厚", 'IF(AND(D{r}="转角",X{r}=1),MAX(G{r}-I{r},0),G{r})', "踢脚高", 'IF(OR(C{r}="吊柜",D{r}="空位"),0,1)', "踢脚板", "柜宽×踢脚高（吊柜/空位无）；L 柜取前脸净长 W1−D"],
    ["踢脚板", "侧板厚", 'IF(AND(D{r}="转角",X{r}=1),MAX(L{r}-I{r},0),0)', "踢脚高", 'IF(AND(D{r}="转角",X{r}=1,C{r}<>"吊柜",L{r}-I{r}>1),1,0)', "踢脚板", "L 柜第二条前脸 W2−D"],
    ["门板", "门板厚", 'IF(E{r}="对开",(G{r}-门缝)/2,G{r}-门缝)', 'H{r}-门缝', 'IF(AND(E{r}<>"无门",D{r}<>"抽屉",LEFT(D{r},2)<>"分段",NOT(AND(D{r}="转角",X{r}=1))),IF(E{r}="对开",2,1),0)', "门板", "普通柜/转角直柜：单门一块、对开两扇各半（逐扇出明细）；L 柜走下面三行"],
    ["门板", "门板厚", "MAX(G{r}-门缝,0)", "H{r}-门缝", 'IF(AND(D{r}="转角",X{r}=1,E{r}<>"无门"),1,0)', "门板", "L 柜 A 段外露部分"],
    ["门板", "门板厚", "MAX(L{r}-门缝,0)", "H{r}-门缝", 'IF(AND(D{r}="转角",X{r}=1,E{r}<>"无门"),1,0)', "门板", "L 柜 B 段外露部分"],
    ["见光板", "见光板厚", "I{r}", "H{r}", "F{r}", "见光板", "柜深×柜高×见光面数（侧板外露面，不含门板）"],
    ["接顶板", "见光板厚", "G{r}", "Z{r}", 'IF(AND(C{r}="吊柜",Z{r}>0),1,0)', "接顶板", "吊柜顶部到天花的封板：柜宽×接顶高度（Z 列，0=不要）"],
    ["见光板-吊柜底", "见光板厚", "G{r}", "I{r}", 'IF(C{r}="吊柜",1,0)', "见光板-吊柜底", "吊柜固定一块底部见光板（柜宽×柜深，仰视面）"],
    ["见光板", "见光板厚", 'IF(E{r}="对开",(G{r}-门缝)/2,G{r}-门缝)', 'H{r}-门缝', 'IF(AND(D{r}="转角",X{r}=0),V{r}+W{r},0)', "见光板", "转角直柜：逐扇指定要见光的门板（清单 V/W 两列，各 0/1）"],
    ["抽侧帮", "抽侧帮厚", f("I{r}-背板厚-抽屉轨道预留"), f("((H{r}-IF(AND(Y{r}=0,C{r}<>\"吊柜\"),0,顶板厚)-底板厚)-(K{r}+1)*门缝)/MAX(K{r},1)-抽体高扣减"), 'IF(D{r}="抽屉",2*K{r}*(R{r}+1),0)', "抽侧帮", "抽深×抽体高×2×每格抽屉数×格数"],
    ["抽底板", "抽底板厚", f(BAYW + "-2*抽屉单侧抽隙-2*抽侧帮厚"), f("I{r}-背板厚-抽屉轨道预留"), 'IF(D{r}="抽屉",K{r}*(R{r}+1),0)', "抽底板", "(格宽-2×抽隙-2×帮厚)×抽深×抽屉数"],
    ["抽后板", "抽后板厚", f(BAYW + "-2*抽屉单侧抽隙-2*抽侧帮厚"), f("((H{r}-IF(AND(Y{r}=0,C{r}<>\"吊柜\"),0,顶板厚)-底板厚)-(K{r}+1)*门缝)/MAX(K{r},1)-抽体高扣减"), 'IF(D{r}="抽屉",K{r}*(R{r}+1),0)', "抽后板", "(格宽-2×抽隙-2×帮厚)×抽体高×抽屉数"],
    ["门板", "门板厚", f(BAYW + "-2*门缝"), f("((H{r}-IF(AND(Y{r}=0,C{r}<>\"吊柜\"),0,顶板厚)-底板厚)-(K{r}+1)*门缝)/MAX(K{r},1)"), 'IF(D{r}="抽屉",K{r}*(R{r}+1),0)', "门板", "抽屉正面=抽面板，按面板类型取名/板厚/单价（与门板同一档）"],
  ];
  const NAME_OF = { side:"侧板", top:"顶板", bottom:"底板", back:"背板", shelf:"层板", vdiv:"竖板",
    kick:"踢脚板", vis:"见光板", dhelp:"抽侧帮", dbox:"抽底板", dback:"抽后板" };
  const LIST_COLS = "CDEFGHIJKLMNOPQRSTUVWXYZ";
  const link = (tpl, row) => tpl.replace(/(?<![A-Za-z!$])([A-Z])\{r\}/g,
    (m, col) => LIST_COLS.includes(col) ? "柜体清单!$" + col + row : m);

  const NOTES = [
    ["面积", "板件宽 × 板件高 × 块数 ÷ 1,000,000（㎡）"],
    ["内宽/内深/内高", "柜宽−2×侧板厚 ／ 柜深−背板厚 ／ 柜高−（本柜顶板厚，清单 Y 列=0 时按 0）−底板厚"],
    ["骨架", "顶底板通长通深 = 柜宽×柜深；侧板夹在中间，高=柜高−顶板厚−底板厚（转角 L 柜通高、空位仍为柜深×柜高）"],
    ["分段柜（上开放格/下抽屉等）", "段数可变，Excel 无法逐段展开 → 该柜在「逐柜明细」里按当前段设置写成固定值（算式说明列标注），其它柜仍是公式；板材汇总/合计走 SUMIFS/SUMIF，照样是公式。改了某柜的段参数，请重新导出一次"],
    ["顶板开关", "每柜单独设（清单 Y 列，1=要 0=不要，默认要）。置 0 时：不出顶板块，且侧板/背板/竖板通到顶，高=柜高−底板厚；吊柜恒有顶板，空位无顶板"],
    ["侧板", "柜深 ×（柜高−顶−底）× 2"],
    ["顶/底板", "柜宽 × 柜深，各 1 块；空位(水槽/灶/洗碗机)为 0 块；地柜（落地柜/高柜）顶板按本柜 Y 列可不要"],
    ["吊柜-底部见光板", "每台吊柜固定 1 块：柜宽 × 柜深（仰视面），走见光板单价"],
    ["吊柜-接顶板", "每柜可选：柜宽 × 接顶高度（清单 Z 列，0=不要），走见光板单价"],
    ["竖板", "（柜深−背板厚）× 内高 × 竖板数（柜体清单 R 列，每柜可填）"],
    ["格宽", "（柜宽−2×侧板厚−竖板数×竖板厚）÷（竖板数+1）"],
    ["背板", "（柜宽−2×侧板厚）×（柜高−顶−底）"],
    ["层板", "普通柜：格宽 ×（柜深−背板厚）× 每格层板数 × 格数；L 柜：每层整块 W1×W2；拉篮柜：按参数「拉篮柜加层板」0/1，默认不加"],
    ["踢脚板", "柜宽 × 踢脚高；吊柜、空位为 0；L 柜 = 两条前脸净长（W1−柜深）+（W2−柜深）"],
    ["门板", "（柜宽−门缝）×（柜高−门缝）；高柜与落地柜同口径（整块通高）；无门为 0；抽屉柜=0（正面逐块计抽面板）"],
    ["门板类型/最低计面积", "每柜在清单 S 列选 平板/造型/玻璃；该柜门板面积合计不足门槛时按门槛计：平板 0.15 ㎡、造型 0.3 ㎡、玻璃 0.5 ㎡（参数表可改）。板材汇总门板三档面积取清单 T 列（已含最低面积），逐柜明细仍是真实块尺寸"],
    ["门板单价", "平板/造型/玻璃各一档单价（参数表单价段），不能合并成一个门板价"],
    ["见光板", "柜深 × 柜高 × 见光面数（1 或 2 面，不含门板）"],
    ["抽屉", "抽体宽=格宽−2×抽隙；抽深=柜深−背板厚−轨道预留；面板高=(柜高−(n+1)×门缝)/n；抽体高=面板高−扣减；侧帮 2n 块、底板/后板/面板各 n 块，n=每格抽屉数×格数"],
    ["转角-L 柜", "矩形掏口：矩形 W1×W2，W2=回转长（清单 L 列）=矩形深度，掏口当废料按整块料算"],
    ["转角-L 柜顶/底板", "各 1 块：W1 × W2（整块）；底板同上"],
    ["转角-L 柜背板", "2 片：W1×柜高 + W2×柜高（各沿一面通高）；参数「转角背板片数」=1 时改为 1 片（W1+W2）×柜高"],
    ["转角-L 柜内侧板", "不出（口径：不用掏口内侧板在中间支撑），清单 M 列恒 0"],
    ["转角-直角柜", "按普通柜口径算骨架/层板/竖板，对开门逐扇指定见光（清单 V/W 列）；层板数可为 0"],
    ["门铰-拉篮", "拉篮=有门板+有轨道、门铰恒 0（与抽屉一致：有轨道不装门铰）"],
    ["门铰", "每扇：门高（=柜高−缝）≤900 用 2 个；超过后每 300 加 1 个，单扇最多 6 个。柜门铰 = 门扇数 × 每扇铰数（对开=2 扇、转角=A/B 段各自按门板数）；翻门按「门铰·每扇翻门」；无门=0；**抽屉柜=0（抽屉没有门，面板走抽面板）**；每柜可手动覆盖。**分段柜按每段自己的门板与段高分别算**。门槛/步长/封顶/单价都在参数表"],
    ["轨道", "抽屉数 × 轨道·每个抽屉；拉篮 1 副 × 轨道·每个拉篮；其余构型 0；每柜可手动覆盖"],
  ];

  const S2 = 2, S3 = 3, S4 = 4, S5 = 5, S6 = 6, S7 = 7;   // xlsx_lite 的样式号

  function cabRows() {
    return S.map((c, i) => {
      const fi = c.form === "吊柜" ? 2 : (c.form === "高柜" ? 3 : 1);
      const segs = Array.isArray(c.segs) ? c.segs : [];
      const confTxt = segs.length ? ("分段:" + segs.map(g => g.conf || "开放格").join("+")) : c.conf_;
      return [i + 1, c.label || ("柜" + (i + 1)), c.form, confTxt, segs.length ? "见各段" : c.doors, +c.vis || 0,
        +c.w, V("H" + fi), V("D" + fi),
        (c.conf_ === "层板" || c.conf_ === "转角") ? (+c.shelves || 0) : 0,   // 转角（L/直角）也有层板
        c.conf_ === "抽屉" ? (+c.drawers || 0) : 0,
        +c.leg2 || 0,
        0,                                        // 转角内侧板已取消（用户：不用中间支撑）
        hingeOf(c), railOf(c),
        (c.conf_ === "空位" || (c.conf_ === "转角" && (c.ct || "L") === "L")) ? 0 : Math.max(+c.vdiv || 0, 0),
        c.dt || "平板", c.ct || "L", +c.dv1 || 0, +c.dv2 || 0,
        c.form === "吊柜" ? 1 : (c.top === 0 ? 0 : 1),     // Y 列：本柜顶板 1=要 0=不要
        c.form === "吊柜" && c.cfTop === 1 ? (+c.cfH || 0) : 0];   // Z 列：吊柜接顶高度 0=不要
    });
  }

  function paramSheet(definedNames) {
    const rows = [[{ v: "参数表（改 C 列即全表重算）", s: S7 }], [], ["序号", "参数", "值", "单位", "说明"].map(v => ({ v, s: 1 }))];
    const nameRow = {}, priceRow = {};
    let r = 4;
    PARAMS.forEach((p, i) => {
      rows.push([i + 1, { v: p[0], s: 2 }, { v: +p[1](), s: 2 }, { v: p[2], s: 2 }, { v: p[3], s: 2 }]);
      nameRow[p[0]] = r; r++;
    });
    rows.push([]);
    rows.push([{ v: "单价（元/㎡，填了才出金额）", s: S7 }]);
    rows.push(["序号", "板件", "单价", "", ""].map(v => ({ v, s: 1 })));
    r += 3;
    UNIT_PRICE.forEach((nm, i) => {
      rows.push([i + 1, { v: nm, s: 2 }, { v: +pv(PRICE_ID[nm]).price || 0, s: 2 }, "", ""]);
      priceRow[nm] = r; r++;
    });
    Object.keys(nameRow).forEach(k => definedNames[dnKey(k)] = "参数!$C$" + nameRow[k]);
    Object.keys(priceRow).forEach(k => definedNames["单价_" + dnKey(k)] = "参数!$C$" + priceRow[k]);
    return { rows, nameRow, priceRow };
  }

  // ================= 公式版 =================
  function buildLive() {
    const dn = {};
    const ps = paramSheet(dn);
    const list = cabRows();
    const listRows = [["编号", "位号/标注", "形态", "构型", "门板", "见光面", "柜宽W", "柜高H", "柜深D",
      "层板数", "抽屉数", "回转长W2", "转角内侧板块数", "门铰数", "轨道数", "B段总长(算)", "转角方格C(算)", "竖板数", "门板类型", "门板计面积(算)", "转角型式", "门板见光1扇", "门板见光2扇", "是否L柜(算)", "顶板(1要0不要)", "接顶高度"].map(v => ({ v, s: 1 }))];
    list.forEach((row, i) => {
      const r = i + 2;
      const dsum = `SUMIFS(逐柜明细!$G:$G,逐柜明细!$A:$A,A${r},逐柜明细!$B:$B,"门板*")`;
      const dmin = `IF(S${r}="造型",门板最低面积造型,IF(S${r}="玻璃",门板最低面积玻璃,门板最低面积平板))`;
      listRows.push(row.slice(0, 15).map(v => ({ v, s: S2 })).concat([
        { f: `IF(D${r}<>"转角",0,IF(转角W2口径=1,MAX(L${r},I${r}),I${r}+L${r}))`, s: S2 },
        { f: `IF(D${r}<>"转角",0,MIN(I${r},G${r}))`, s: S2 },
        { v: row[15], s: S2 },
        { v: row[16], s: S2 },
        { f: `IF(${dsum}=0,0,MAX(${dsum},${dmin}))`, s: S3 },
        { v: row[17], s: S2 },
        { v: row[18], s: S2 },
        { v: row[19], s: S2 },
        { f: `IF(AND(D${r}="转角",U${r}="L"),1,0)`, s: S2 },
        { v: row[20], s: S2 }, { v: row[21], s: S2 }]));
    });

    const det = [["柜号", "板件", "板厚", "板件宽", "板件高", "块数", "面积㎡", "单价元/㎡", "金额元", "算式说明"].map(v => ({ v, s: 1 }))];
    let ro = 2;
    list.forEach((row, li) => {
      const idx = li + 2;                       // 柜体清单里的行号（表头在第 1 行）
      const cabSeg = S[li];
      if (Array.isArray(cabSeg.segs) && cabSeg.segs.length) {
        // 分段柜（上开放格/下抽屉…）段数不固定，Excel 无法逐段展开 → 明细按当前参数写成固定值；
        // 柜体清单的 T 列（门板计面积）仍走 SUMIFS，所以板材汇总/合计照样是公式、照样汇总。
        // 用户口径：分段柜只到"整柜"粒度 —— 按板件名合并，不细到第几个抽屉/层板
        const grp = new Map();
        parts(cabSeg).forEach(pt => {
          const isDoor = String(pt.id).startsWith("door_");
          const nm = isDoor ? ("门板-" + (cabSeg.dt || "平板")) : (pt.nm || NAME_OF[pt.id] || pt.id);
          const key = nm;
          const g = grp.get(key) || { nm, q: 0, area: 0, ws: new Set(), hs: new Set(), door: isDoor };
          g.q += +pt.q; g.area += +pt.area; g.ws.add(+pt.w); g.hs.add(+pt.h);
          grp.set(key, g);
        });
        grp.forEach(g => {
          const priceKey = g.door ? g.nm : g.nm;
          const th = g.door ? +pv("door_" + ({ "平板": "p", "造型": "m", "玻璃": "g" }[cabSeg.dt || "平板"] || "p")).t
                            : +pv(Object.keys(NAME_OF).find(k => NAME_OF[k] === g.nm) || "side").t;
          det.push([{ v: row[0], s: S2 }, { v: g.nm, s: S2 }, { v: th, s: S2 },
            { v: g.ws.size === 1 ? [...g.ws][0] : "混合", s: S2 },
            { v: g.hs.size === 1 ? [...g.hs][0] : "混合", s: S2 },
            { v: g.q, s: S2 }, { v: +g.area.toFixed(6), s: S3 },
            { f: "单价_" + dnKey(priceKey), s: S4 }, { f: `G${ro}*H${ro}`, s: S4 },
            { v: "分段柜：整柜按板件合并（改段参数后重新导出）", s: S2 }]);
          ro++;
        });
        return;
      }
      BOARDS.forEach(b => {
        const isDoor = String(b[5]).startsWith("门板");
        const dtRef = `柜体清单!$S$${idx}`;
        det.push([{ v: row[0], s: S2 },
          isDoor ? { f: `="门板-"&${dtRef}`, s: S2 } : { v: b[0], s: S2 },
          isDoor ? { f: `IF(${dtRef}="造型",门板造型厚,IF(${dtRef}="玻璃",门板玻璃厚,门板平板厚))`, s: S2 }
                 : { f: b[1], s: S2 },
          { f: link(b[2], idx), s: S2 }, { f: link(b[3], idx), s: S2 }, { f: link(b[4], idx), s: S2 },
          { f: `IF(F${ro}>0,D${ro}*E${ro}*F${ro}/1000000,0)`, s: S3 },
          isDoor ? { f: `IF(${dtRef}="造型",单价_门板造型,IF(${dtRef}="玻璃",单价_门板玻璃,单价_门板平板))`, s: S4 }
                 : { f: "单价_" + dnKey(b[5]), s: S4 },
          { f: `G${ro}*H${ro}`, s: S4 }, { v: b[6], s: S2 }]);
        ro++;
      });
    });
    const last = ro - 1;

    const sum = [["板件", "板厚", "面积㎡", "单价元/㎡", "金额元"].map(v => ({ v, s: 1 }))];
    const listLast = list.length + 1;
    UNIT_PRICE.forEach((p, i) => {
      const r = i + 2;
      const isDoor = p.startsWith("门板-");
      const areaF = isDoor
        ? `SUMIFS(柜体清单!$T$2:$T$${listLast},柜体清单!$S$2:$S$${listLast},"${p.slice(3)}")`
        : `SUMIF(逐柜明细!$B$2:$B$${last},A${r},逐柜明细!$G$2:$G$${last})`;
      sum.push([{ v: p, s: S2 },
        { f: `IFERROR(INDEX(逐柜明细!$C$2:$C$${last},MATCH(A${r},逐柜明细!$B$2:$B$${last},0)),"")`, s: S2 },
        { f: areaF, s: S3 },
        { f: `IFERROR(INDEX(逐柜明细!$H$2:$H$${last},MATCH(A${r},逐柜明细!$B$2:$B$${last},0)),"")`, s: S4 },
        { f: `IFERROR(C${r}*D${r},0)`, s: S4 }]);
    });
    const rt = UNIT_PRICE.length + 2;
    sum.push([{ v: "合计", s: S5 }, "", { f: `SUM(C2:C${rt - 1})`, s: S5 }, "", { f: `SUM(E2:E${rt - 1})`, s: S5 }]);

    const noteRows = [["项", "算法（全部写成公式，改参数自动重算）"].map(v => ({ v, s: 1 }))]
      .concat(NOTES.map(n => [{ v: n[0], s: S2 }, { v: n[1], s: S6 }]));

    const n = list.length + 1;
    const hw = [["项", "数量", "单位", "单价元", "金额元"].map(v => ({ v, s: 1 })),
      [{ v: "门铰", s: S2 }, { f: `SUM(柜体清单!N2:N${n})`, s: S2 }, { v: "件", s: S2 }, { f: "门铰单价", s: S4 }, { f: "B2*D2", s: S4 }],
      [{ v: "轨道", s: S2 }, { f: `SUM(柜体清单!O2:O${n})`, s: S2 }, { v: "副", s: S2 }, { f: "轨道单价", s: S4 }, { f: "B3*D3", s: S4 }],
      [{ v: "合计", s: S5 }, "", "", "", { f: "SUM(E2:E3)", s: S5 }]];
    return XLSXLite.build([
      { name: "参数", cols: [6, 20, 12, 8, 46], rows: ps.rows },
      { name: "柜体清单", cols: [6, 12, 9, 8, 9, 8, 9, 8, 8, 7, 7, 9, 12, 8, 8, 11, 12, 8, 9, 12, 9, 9, 9, 10], rows: listRows, freeze: 1 },
      { name: "逐柜明细", cols: [6, 16, 8, 10, 10, 7, 10, 10, 10, 40], rows: det, freeze: 1 },
      { name: "板材汇总", cols: [16, 8, 12, 12, 12], rows: sum },
      { name: "五金汇总", cols: [10, 10, 8, 10, 12], rows: hw },
      { name: "算法说明", cols: [20, 110], rows: noteRows },
    ], { definedNames: dn });
  }

  // ================= 死值快照 =================
  function buildFixed() {
    const dn = {};
    const ps = paramSheet(dn);
    const list = cabRows();
    const listRows = [["编号", "位号/标注", "形态", "构型", "门板", "见光面", "柜宽W", "柜高H", "柜深D",
      "层板数", "抽屉数", "回转长W2", "转角内侧板块数", "门铰数", "轨道数", "B段总长(算)", "转角方格C(算)", "竖板数", "门板类型", "门板计面积(算)", "转角型式", "门板见光1扇", "门板见光2扇", "是否L柜(算)", "顶板(1要0不要)", "接顶高度"].map(v => ({ v, s: 1 }))];
    list.forEach((row, i) => {
      const D = +row[8], W = +row[6], W2 = +row[11];
      const isC = row[3] === "转角";
      const B = isC ? (+V("cnW2Ref") === 1 ? Math.max(W2, D) : D + W2) : 0;
      const C = isC ? Math.min(D, W) : 0;
      const cab = S[i], doorReal = parts(cab).filter(p => !p.min && p.id.startsWith("door")).reduce((a, p) => a + p.area, 0);
      const dtk = row[16], dmin = +V({ "平板": "dmP", "造型": "dmM", "玻璃": "dmG" }[dtk] || "dmP") || 0;
      const doorBill = doorReal <= 0 ? 0 : Math.max(doorReal, dmin);
      listRows.push(row.slice(0, 15).map(v => ({ v, s: S2 })).concat([{ v: +B.toFixed(1), s: S2 }, { v: +C.toFixed(1), s: S2 },
        { v: row[15], s: S2 }, { v: row[16], s: S2 }, { v: +doorBill.toFixed(4), s: S3 },
        { v: row[17], s: S2 }, { v: row[18], s: S2 }, { v: row[19], s: S2 },
        { v: (cab.conf_ === "转角" && (cab.ct || "L") === "L") ? 1 : 0, s: S2 }, { v: row[20], s: S2 }, { v: row[21], s: S2 }]));
    });

    const det = [["柜号", "板件", "板厚", "板件宽", "板件高", "块数", "面积㎡", "单价元/㎡", "金额元", "算式说明"].map(v => ({ v, s: 1 }))];
    const agg = {};
    S.forEach((c, i) => {
      parts(c).forEach(p => {
        const nm = p.nm || (ST.part.find(q => q.id === p.id) || { name: p.id }).name, t = +pv(p.id).t, pr = +pv(p.id).price || 0;
        det.push([{ v: i + 1, s: S2 }, { v: nm, s: S2 }, { v: t, s: S2 }, { v: p.w, s: S2 }, { v: p.h, s: S2 },
          { v: p.q, s: S2 }, { v: +p.area.toFixed(4), s: S3 }, { v: pr, s: S4 },
          { v: +(p.area * pr).toFixed(2), s: S4 }, { v: `${p.w}×${p.h}×${p.q}${p.role ? " " + p.role : ""}`, s: S2 }]);
        const k = nm + "@" + t;
        agg[k] = agg[k] || { nm, t, pr, a: 0 };
        agg[k].a += p.area;
      });
    });

    const order = UNIT_PRICE.slice();
    const keys = Object.keys(agg).sort((x, y) => order.indexOf(agg[x].nm) - order.indexOf(agg[y].nm));
    const sum = [["板件", "板厚", "面积㎡", "单价元/㎡", "金额元"].map(v => ({ v, s: 1 }))];
    let ta = 0, tm = 0;
    keys.forEach(k => {
      const g = agg[k]; ta += g.a; tm += g.a * g.pr;
      sum.push([{ v: g.nm, s: S2 }, { v: g.t, s: S2 }, { v: +g.a.toFixed(4), s: S3 },
        { v: g.pr, s: S4 }, { v: +(g.a * g.pr).toFixed(2), s: S4 }]);
    });
    sum.push([{ v: "合计", s: S5 }, "", { v: +ta.toFixed(4), s: S5 }, "", { v: +tm.toFixed(2), s: S5 }]);

    const noteRows = [["项", "算法"].map(v => ({ v, s: 1 }))]
      .concat(NOTES.map(n => [{ v: n[0], s: S2 }, { v: n[1], s: S6 }]));
    let hq = 0, rq = 0; S.forEach(c => { hq += hingeOf(c); rq += railOf(c); });
    const hp = +V("hjPrice") || 0, rp = +V("rlPrice") || 0;
    const hw = [["项", "数量", "单位", "单价元", "金额元"].map(v => ({ v, s: 1 })),
      [{ v: "门铰", s: S2 }, { v: hq, s: S2 }, { v: "件", s: S2 }, { v: hp, s: S4 }, { v: +(hq * hp).toFixed(2), s: S4 }],
      [{ v: "轨道", s: S2 }, { v: rq, s: S2 }, { v: "副", s: S2 }, { v: rp, s: S4 }, { v: +(rq * rp).toFixed(2), s: S4 }],
      [{ v: "合计", s: S5 }, "", "", "", { v: +(hq * hp + rq * rp).toFixed(2), s: S5 }]];
    return XLSXLite.build([
      { name: "参数", cols: [6, 20, 12, 8, 46], rows: ps.rows },
      { name: "柜体清单", cols: [6, 12, 9, 8, 9, 8, 9, 8, 8, 7, 7, 9, 12, 8, 8, 11, 12, 8, 9, 12, 9, 9, 9, 10], rows: listRows, freeze: 1 },
      { name: "逐柜明细", cols: [6, 16, 8, 10, 10, 7, 10, 10, 10, 40], rows: det, freeze: 1 },
      { name: "板材汇总", cols: [16, 8, 12, 12, 12], rows: sum },
      { name: "五金汇总", cols: [10, 10, 8, 10, 12], rows: hw },
      { name: "算法说明", cols: [20, 110], rows: noteRows },
    ], { definedNames: dn });
  }

  function save(u8, filename) {
    const b = new Blob([u8], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b); a.download = filename; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function guard() {
    if (typeof XLSXLite === "undefined") { alert("缺少 xlsx_lite.js，无法导出"); return false; }
    return true;
  }
  window.exportLive = function () {
    if (!guard()) return;
    // 分段柜已支持：其明细按当前段设置写固定值（段数可变，Excel 无法逐段展开），
    // 板材汇总/合计仍是公式汇总，改其它柜参数照样重算。
    try { save(buildLive(), "橱柜算料_公式版.xlsx"); }
    catch (e) { alert("导出失败：" + e.message); }
  };
  window.exportExcel = function () {
    if (!guard()) return;
    try { save(buildFixed(), "橱柜算料.xlsx"); }
    catch (e) { alert("导出失败：" + e.message); }
  };
  window.__cabXlsx = { buildLive, buildFixed };   // 供自测
})();
