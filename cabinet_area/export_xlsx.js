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
    ["背板厚", () => pv("back").t, "mm", "薄板"],
    ["见光板厚", () => pv("vis").t, "mm", "外露装饰面"],
    ["抽侧帮厚", () => pv("dhelp").t, "mm", "抽屉"],
    ["抽底板厚", () => pv("dbox").t, "mm", ""],
    ["抽后板厚", () => pv("dback").t, "mm", ""],
    ["抽面板厚", () => pv("dface").t, "mm", ""],
    ["门板厚", () => pv("door").t, "mm", "门板基材"],
    ["踢脚高", () => V("K"), "mm", "吊柜无踢脚"],
    ["门缝", () => V("G"), "mm", "四周留缝"],
    ["抽屉单侧抽隙", () => V("dGap"), "mm", "三节轨常规"],
    ["抽屉轨道预留", () => V("dRail"), "mm", "抽深=柜深-背板-此项"],
    ["抽体高扣减", () => V("dCut"), "mm", "抽体高=面板高-此项"],
    ["转角背板片数", () => V("cnBack"), "片", "2=每段各一片(默认)"],
    ["转角背板重叠扣减", () => V("cnOver"), "mm", "两片在转角处扣一个板厚"],
    ["转角W2口径", () => V("cnW2Ref"), "0/1", "0=W2从转角方格外端量(净长) 1=从相邻墙内角量(含方格)"],
    ["转角方格门", () => V("cnDoor"), "0/1", "0=盲角无门 1=方格装门"],
  ];
  const UNIT_PRICE = ["侧板", "顶板", "底板", "层板", "背板", "门板", "见光板", "踢脚板", "转角内侧板",
    "抽侧帮", "抽底板", "抽后板", "抽面板"];
  const PRICE_ID = { "侧板": "side", "顶板": "top", "底板": "bottom", "层板": "shelf", "背板": "back",
    "门板": "door", "见光板": "vis", "踢脚板": "kick", "转角内侧板": "corner",
    "抽侧帮": "dhelp", "抽底板": "dbox", "抽后板": "dback", "抽面板": "dface" };

  const f = (tpl, min) => `MAX(${tpl},${min == null ? 0 : min})`;   // 尺寸夹到 >=0
  // ---------- 明细模板（与 make_excel.py 的 BOARDS 同口径）----------
  const BOARDS = [
    ["侧板", "侧板厚", "I{r}", "H{r}", "2", "侧板", "柜深×柜高×2（侧板包住顶底板）"],
    ["转角内侧板", "侧板厚", "I{r}", "H{r}", "M{r}", "转角内侧板", "转角共用/各段内侧竖板，块数见清单M列"],
    ["顶板", "顶板厚", 'IF(D{r}="转角",G{r}+N{r}-O{r},G{r}-2*侧板厚)', "I{r}", 'IF(D{r}="空位",0,1)', "顶板", "普通柜=柜宽-2×侧板厚；转角=W1+B-方格"],
    ["底板", "底板厚", 'IF(D{r}="转角",G{r}+N{r}-O{r},G{r}-2*侧板厚)', "I{r}", 'IF(D{r}="空位",0,1)', "底板", "同上"],
    ["背板", "背板厚", 'IF(D{r}="转角",IF(转角背板片数=2,G{r},G{r}+N{r}-O{r}),G{r}-2*侧板厚)', 'H{r}-顶板厚-底板厚', "1", "背板", "转角两片时：本段(W1)一片"],
    ["背板", "背板厚", 'IF(D{r}="转角",MAX(N{r}-转角背板重叠扣减,0),0)', 'H{r}-顶板厚-底板厚', 'IF(AND(D{r}="转角",转角背板片数=2),1,0)', "背板", "转角第二片（B段，扣重叠）"],
    ["层板", "层板厚", 'IF(D{r}="转角",0,G{r}-2*侧板厚)', "I{r}-背板厚", 'IF(D{r}="层板",J{r},IF(D{r}="拉篮",1,0))', "层板", "内宽×(柜深-背板厚)×层板数"],
    ["踢脚板", "侧板厚", 'IF(D{r}="转角",G{r}+N{r}-O{r},G{r})', "踢脚高", 'IF(OR(C{r}="吊柜",D{r}="空位"),0,1)', "踢脚板", "柜宽×踢脚高（吊柜/空位无）"],
    ["门板", "门板厚", "G{r}-门缝", 'IF(C{r}="高柜",H{r}/2,H{r})-门缝', 'IF(AND(E{r}<>"无门",D{r}<>"转角"),1,0)', "门板", "（柜宽-缝）×（柜高-缝）；高柜按上下两段"],
    ["门板", "门板厚", "MAX(G{r}-O{r}-门缝,0)", "H{r}-门缝", 'IF(AND(E{r}<>"无门",D{r}="转角"),1,0)', "门板", "转角A段外露部分"],
    ["门板", "门板厚", "MAX(N{r}-O{r}-门缝,0)", "H{r}-门缝", 'IF(AND(E{r}<>"无门",D{r}="转角"),1,0)', "门板", "转角B段外露部分"],
    ["门板", "门板厚", f("O{r}-门缝",1), "H{r}-门缝", 'IF(AND(D{r}="转角",转角方格门=1,E{r}<>"无门"),1,0)', "门板", "仅当转角方格装门"],
    ["见光板", "见光板厚", "I{r}", "H{r}", "F{r}", "见光板", "柜深×柜高×见光面数（不含门板）"],
    ["抽侧帮", "抽侧帮厚", f("I{r}-背板厚-抽屉轨道预留"), f("(H{r}-(K{r}+1)*门缝)/MAX(K{r},1)-抽体高扣减"), 'IF(D{r}="抽屉",2*K{r},0)', "抽侧帮", "抽深×抽体高×2×抽屉数"],
    ["抽底板", "抽底板厚", f("G{r}-2*侧板厚-2*抽屉单侧抽隙-2*抽侧帮厚"), f("I{r}-背板厚-抽屉轨道预留"), 'IF(D{r}="抽屉",K{r},0)', "抽底板", "(抽体宽-2×帮厚)×抽深"],
    ["抽后板", "抽后板厚", f("G{r}-2*侧板厚-2*抽屉单侧抽隙-2*抽侧帮厚"), f("(H{r}-(K{r}+1)*门缝)/MAX(K{r},1)-抽体高扣减"), 'IF(D{r}="抽屉",K{r},0)', "抽后板", "(抽体宽-2×帮厚)×抽体高"],
    ["抽面板", "抽面板厚", f("G{r}-2*门缝"), f("(H{r}-(K{r}+1)*门缝)/MAX(K{r},1)"), 'IF(D{r}="抽屉",K{r},0)', "抽面板", "（柜宽-2缝）×面板高×抽屉数"],
  ];
  const LIST_COLS = "CDEFGHIJKLMNO";
  const link = (tpl, row) => tpl.replace(/(?<![A-Za-z!$])([A-Z])\{r\}/g,
    (m, col) => LIST_COLS.includes(col) ? "柜体清单!$" + col + row : m);

  const NOTES = [
    ["面积", "板件宽 × 板件高 × 块数 ÷ 1,000,000（㎡）"],
    ["内宽/内深/内高", "柜宽−2×侧板厚 ／ 柜深−背板厚 ／ 柜高−顶板厚−底板厚"],
    ["侧板", "柜深 × 柜高 × 2"],
    ["顶/底板", "（柜宽−2×侧板厚）× 柜深，各 1 块；空位(水槽/灶/洗碗机)为 0 块"],
    ["背板", "（柜宽−2×侧板厚）×（柜高−顶−底）"],
    ["层板", "（柜宽−2×侧板厚）×（柜深−背板厚）× 层板数"],
    ["踢脚板", "柜宽 × 踢脚高；吊柜、空位为 0"],
    ["门板", "（柜宽−门缝）×（柜高−门缝）；高柜按上下两段，每段高=柜高/2；无门为 0"],
    ["见光板", "柜深 × 柜高 × 见光面数（1 或 2 面，不含门板）"],
    ["抽屉", "抽体宽=柜宽−2×侧板厚−2×抽隙；抽深=柜深−背板厚−轨道预留；面板高=(柜高−(n+1)×门缝)/n；抽体高=面板高−扣减；侧帮 2n 块、底板/后板/面板各 n 块"],
    ["转角-B段总长", "口径0：B=柜深+W2；口径1：B=W2（清单 N 列自动算）"],
    ["转角-顶/底板", "（W1+B−转角方格）× 柜深，各 1 块"],
    ["转角-背板", "2 片：W1×内高 +（B−重叠扣减）×内高；1 片：（W1+B−方格）×内高"],
    ["转角-内侧板", "柜深 × 柜高 × 块数（清单 M 列）"],
    ["转角-门板", "两段外露：（W1−C−缝）×（柜高−缝）+（B−C−缝）×（柜高−缝）；方格装门时再加（C−缝）×（柜高−缝）"],
  ];

  const S2 = 2, S3 = 3, S4 = 4, S5 = 5, S6 = 6, S7 = 7;   // xlsx_lite 的样式号

  function cabRows() {
    return S.map((c, i) => {
      const fi = c.form === "吊柜" ? 2 : (c.form === "高柜" ? 3 : 1);
      return [i + 1, c.label || ("柜" + (i + 1)), c.form, c.conf_, c.doors, +c.vis || 0,
        +c.w, V("H" + fi), V("D" + fi),
        c.conf_ === "层板" ? (+c.shelves || 0) : 0,
        c.conf_ === "抽屉" ? (+c.drawers || 0) : 0,
        +c.leg2 || 0,
        c.conf_ === "转角" ? (+V("cnSide") || 1) : 0];
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
    Object.keys(nameRow).forEach(k => definedNames[k] = "参数!$C$" + nameRow[k]);
    Object.keys(priceRow).forEach(k => definedNames["单价_" + k] = "参数!$C$" + priceRow[k]);
    return { rows, nameRow, priceRow };
  }

  // ================= 公式版 =================
  function buildLive() {
    const dn = {};
    const ps = paramSheet(dn);
    const list = cabRows();
    const listRows = [["编号", "位号/标注", "形态", "构型", "门板", "见光面", "柜宽W", "柜高H", "柜深D",
      "层板数", "抽屉数", "回转长W2", "转角内侧板块数", "B段总长(算)", "转角方格C(算)"].map(v => ({ v, s: 1 }))];
    list.forEach((row, i) => {
      const r = i + 2;
      listRows.push(row.map(v => ({ v, s: S2 })).concat([
        { f: `IF(D${r}<>"转角",0,IF(转角W2口径=1,MAX(L${r},I${r}),I${r}+L${r}))`, s: S2 },
        { f: `IF(D${r}<>"转角",0,MIN(I${r},G${r}))`, s: S2 }]));
    });

    const det = [["柜号", "板件", "板厚", "板件宽", "板件高", "块数", "面积㎡", "单价元/㎡", "金额元", "算式说明"].map(v => ({ v, s: 1 }))];
    let ro = 2;
    list.forEach((row, li) => {
      const idx = li + 2;                       // 柜体清单里的行号（表头在第 1 行）
      BOARDS.forEach(b => {
        det.push([{ v: row[0], s: S2 }, { v: b[0], s: S2 }, { f: b[1], s: S2 },
          { f: link(b[2], idx), s: S2 }, { f: link(b[3], idx), s: S2 }, { f: link(b[4], idx), s: S2 },
          { f: `IF(F${ro}>0,D${ro}*E${ro}*F${ro}/1000000,0)`, s: S3 },
          { f: "单价_" + b[5], s: S4 }, { f: `G${ro}*H${ro}`, s: S4 }, { v: b[6], s: S2 }]);
        ro++;
      });
    });
    const last = ro - 1;

    const sum = [["板件", "板厚", "面积㎡", "单价元/㎡", "金额元"].map(v => ({ v, s: 1 }))];
    UNIT_PRICE.forEach((p, i) => {
      const r = i + 2;
      sum.push([{ v: p, s: S2 },
        { f: `IFERROR(INDEX(逐柜明细!$C$2:$C$${last},MATCH(A${r},逐柜明细!$B$2:$B$${last},0)),"")`, s: S2 },
        { f: `SUMIF(逐柜明细!$B$2:$B$${last},A${r},逐柜明细!$G$2:$G$${last})`, s: S3 },
        { f: `IFERROR(INDEX(逐柜明细!$H$2:$H$${last},MATCH(A${r},逐柜明细!$B$2:$B$${last},0)),"")`, s: S4 },
        { f: `IFERROR(C${r}*D${r},0)`, s: S4 }]);
    });
    const rt = UNIT_PRICE.length + 2;
    sum.push([{ v: "合计", s: S5 }, "", { f: `SUM(C2:C${rt - 1})`, s: S5 }, "", { f: `SUM(E2:E${rt - 1})`, s: S5 }]);

    const noteRows = [["项", "算法（全部写成公式，改参数自动重算）"].map(v => ({ v, s: 1 }))]
      .concat(NOTES.map(n => [{ v: n[0], s: S2 }, { v: n[1], s: S6 }]));

    return XLSXLite.build([
      { name: "参数", cols: [6, 20, 12, 8, 46], rows: ps.rows },
      { name: "柜体清单", cols: [6, 12, 9, 8, 9, 8, 9, 8, 8, 7, 7, 9, 12, 11, 12], rows: listRows, freeze: 1 },
      { name: "逐柜明细", cols: [6, 16, 8, 10, 10, 7, 10, 10, 10, 40], rows: det, freeze: 1 },
      { name: "板材汇总", cols: [16, 8, 12, 12, 12], rows: sum },
      { name: "算法说明", cols: [20, 110], rows: noteRows },
    ], { definedNames: dn });
  }

  // ================= 死值快照 =================
  function buildFixed() {
    const dn = {};
    const ps = paramSheet(dn);
    const list = cabRows();
    const listRows = [["编号", "位号/标注", "形态", "构型", "门板", "见光面", "柜宽W", "柜高H", "柜深D",
      "层板数", "抽屉数", "回转长W2", "转角内侧板块数", "B段总长(算)", "转角方格C(算)"].map(v => ({ v, s: 1 }))];
    list.forEach((row, i) => {
      const D = +row[8], W = +row[6], W2 = +row[11];
      const isC = row[3] === "转角";
      const B = isC ? (+V("cnW2Ref") === 1 ? Math.max(W2, D) : D + W2) : 0;
      const C = isC ? Math.min(D, W) : 0;
      listRows.push(row.map(v => ({ v, s: S2 })).concat([{ v: +B.toFixed(1), s: S2 }, { v: +C.toFixed(1), s: S2 }]));
    });

    const det = [["柜号", "板件", "板厚", "板件宽", "板件高", "块数", "面积㎡", "单价元/㎡", "金额元", "算式说明"].map(v => ({ v, s: 1 }))];
    const agg = {};
    S.forEach((c, i) => {
      parts(c).forEach(p => {
        const nm = (ST.part.find(q => q.id === p.id) || { name: p.id }).name, t = +pv(p.id).t, pr = +pv(p.id).price || 0;
        det.push([{ v: i + 1, s: S2 }, { v: nm, s: S2 }, { v: t, s: S2 }, { v: p.w, s: S2 }, { v: p.h, s: S2 },
          { v: p.q, s: S2 }, { v: +p.area.toFixed(4), s: S3 }, { v: pr, s: S4 },
          { v: +(p.area * pr).toFixed(2), s: S4 }, { v: `${p.w}×${p.h}×${p.q}`, s: S2 }]);
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
    return XLSXLite.build([
      { name: "参数", cols: [6, 20, 12, 8, 46], rows: ps.rows },
      { name: "柜体清单", cols: [6, 12, 9, 8, 9, 8, 9, 8, 8, 7, 7, 9, 12, 11, 12], rows: listRows, freeze: 1 },
      { name: "逐柜明细", cols: [6, 16, 8, 10, 10, 7, 10, 10, 10, 40], rows: det, freeze: 1 },
      { name: "板材汇总", cols: [16, 8, 12, 12, 12], rows: sum },
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
