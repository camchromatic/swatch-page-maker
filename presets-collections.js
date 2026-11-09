// Collections: several sets of one product line offered as a single option in the
// Set menu. The sheet shows each set as its own group with the set's name in the
// header, and keeps every set complete (colors two sets share appear in both).
// The member sets no longer show on their own in the menu.
(function () {
  function collect(id, name, title, group) {
    const members = PRESETS.filter((p) => p.group === group && !p.collection && !p.hidden);
    if (members.length < 2) return;
    for (const m of members) m.hidden = true;
    PRESETS.splice(PRESETS.indexOf(members[0]), 0, { // in its sets' place, so menu order stays
      id, group, name, title, medium: members[0].medium,
      collection: members.map((m) => m.id),
      codes: members.flatMap((m) => m.codes),
    });
  }

  // A collection from a hand-picked list of (already hidden) presets. forceHeaders: the
  // headers can't be switched off, because the group names are the point of it.
  function collectIds(id, name, title, group, ids, forceHeaders) {
    const members = ids.map((m) => PRESETS.find((p) => p.id === m)).filter(Boolean);
    if (!members.length) return;
    PRESETS.splice(PRESETS.indexOf(members[0]), 0, {
      id, group, name, title, medium: members[0].medium, collection: members.map((m) => m.id),
      codes: members.flatMap((m) => m.codes), forceHeaders,
    });
  }

  collect("faber-pitt-dual-all", "All 10 Dual Marker sets", "Faber-Castell PITT Dual Marker collection",
    "Faber-Castell PITT Artist Pen Dual Marker");
  // Copic acrea: the eight 6-color sets make up all 48 colors exactly.
  collect("copic-acrea-all-sets", "All 48 colors", "Copic acrea collection", "Copic acrea");
  // Posca: one option per tip, grouped by color family with headers.
  for (const [tip, name, fams] of [
    ["pc-1m", "All 36 colors", ["standard", "pastel", "metallic"]],
    ["pc-1mr", "All 22 colors", ["standard", "pastel", "metallic"]],
    ["pc-3m", "All 62 colors", ["standard", "pastel", "metallic", "glitter", "sparkling"]],
    ["pc-5m", "All 68 colors", ["standard", "pastel", "fluorescent", "metallic"]],
    ["pc-7m", "All 15 colors", ["standard"]],
    ["pc-8k", "All 39 colors", ["standard", "pastel", "fluorescent", "metallic"]],
    ["pc-17k", "All 10 colors", ["standard", "metallic"]],
    ["pcf-350", "All 10 colors", ["standard", "metallic"]],
    ["pc-5br", "All 36 colors", ["standard", "pastel"]],
  ]) {
    const group = "Posca " + tip.toUpperCase();
    collectIds(`posca-${tip}-grouped`, name, `${group} ${name}`, group, fams.map((f) => `posca-fam-${tip}-${f}`), true);
  }
  // Mildliner: all 41 colors grouped by their 5-color packs, newest first.
  collectIds("mildliner-all-grouped", "All 41 colors", "Zebra Mildliner All 41 colors", "Zebra Mildliner",
    ["calm", "natural", "cool-refined", "fluorescent", "deep-warm", "friendly", "bright", "gentle"].map((f) => `mildliner-fam-${f}`));
  // Karin: all 73 colors grouped by the six 12-color sets (they hold the 72 colors once
  // each), plus the Colorless Blender.
  for (const [line, slug] of [["Brushmarker PRO", "brushmarker-pro"], ["Realbrush PRO", "realbrush-pro"]]) {
    collectIds(`karin-markers-${slug}-grouped`, "All 73 colors", `Karin ${line} All 73 colors`, `Karin ${line}`,
      ["flower", "sky", "sun-and-tree", "skin", "gray", "neon"].map((f) => `karin-markers-${slug}-${f}-colors-12`).concat(`karin-markers-${slug}-blender`));
  }
})();
