#target photoshop
// Nefertiti label - adds editable text layers to Nefertiti_Varicose_Serum_layers.psd
// Fonts: Marcellus (free, Google Fonts), Myriad Pro + Adobe Arabic (ship with Photoshop)
app.bringToFront();
if (!app.documents.length) {
    var f = File.openDialog("Open Nefertiti_Varicose_Serum_layers.psd", "*.psd");
    if (f) app.open(f);
}
var doc = app.activeDocument;
var oldRuler = app.preferences.rulerUnits, oldType = app.preferences.typeUnits;
app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;
var res = doc.resolution;
if (res != 72) doc.resizeImage(undefined, undefined, 72, ResampleMethod.NONE);

var S = [
["Front Bottle", "Brand", "NEFERTITI", "Marcellus-Regular", 30, [217, 166, 79], 454, 459, "C", 0],
["Front Bottle", "Tagline", "for natural oils and herbs", "MyriadPro-Regular", 11.5, [217, 166, 79], 455, 474, "C", 0],
["Front Bottle", "Title 1", "VARICOSE", "Marcellus-Regular", 37, [238, 232, 220], 447, 529, "C", 0],
["Front Bottle", "Title 2", "VEIN", "Marcellus-Regular", 41, [238, 232, 220], 453, 568, "C", 0],
["Front Bottle", "Title 3", "SERUM", "Marcellus-Regular", 44, [217, 166, 79], 453, 611, "C", 0],
["Front Bottle", "Subtitle", "With Cypress Oil\r& Vitamin E", "MyriadPro-Regular", 19, [238, 232, 220], 455, 640, "C", 23],
["Front Bottle", "Benefit 1a", "Supports", "MyriadPro-Regular", 19, [238, 232, 220], 423, 984, "L", 0],
["Front Bottle", "Benefit 1b", "Healthy Circulation", "MyriadPro-Regular", 16, [238, 232, 220], 423, 1003, "L", 0],
["Front Bottle", "Benefit 2", "Relieves\rHeavy Legs", "MyriadPro-Regular", 19, [238, 232, 220], 427, 1052, "L", 21],
["Front Bottle", "Benefit 3", "Soothes &\rRefreshes", "MyriadPro-Regular", 19, [238, 232, 220], 429, 1123, "L", 20],
["Front Bottle", "Volume", "100 ml", "MyriadPro-Regular", 25, [238, 232, 220], 403, 1205, "L", 0],
["Back Bottle", "Brand", "NEFERTITI", "Marcellus-Regular", 21, [217, 166, 79], 908, 432, "C", 0],
["Back Bottle", "Tagline", "for natural oils and herbs", "MyriadPro-Regular", 9.5, [217, 166, 79], 908, 444, "C", 0],
["Back Bottle", "Description EN", "A natural formula enriched\rwith Cypress Oil and Vitamin E\rthat helps improve blood\rcirculation, reduce the appearance\rof varicose veins and relieve\rheaviness and fatigue in the legs.", "MyriadPro-Regular", 14.5, [238, 232, 220], 907, 478, "C", 16.2],
["Back Bottle", "Description AR", "\u062a\u0631\u0643\u064a\u0628\u0629 \u0637\u0628\u064a\u0639\u064a\u0629 \u0628\u0632\u064a\u062a \u0627\u0644\u0633\u0631\u0648\r\u0648\u0641\u064a\u062a\u0627\u0645\u064a\u0646 E \u062a\u0633\u0627\u0639\u062f \u0639\u0644\u0649 \u062a\u062d\u0633\u064a\u0646\r\u0627\u0644\u062f\u0648\u0631\u0629 \u0627\u0644\u062f\u0645\u0648\u064a\u0629\u060c \u0648\u062a\u0642\u0644\u064a\u0644 \u0645\u0638\u0647\u0631\r\u0627\u0644\u062f\u0648\u0627\u0644\u064a\u060c \u0648\u062a\u062e\u0641\u064a\u0641 \u0627\u0644\u0634\u0639\u0648\u0631\r\u0628\u0627\u0644\u062b\u0642\u0644 \u0648\u0627\u0644\u062a\u0639\u0628 \u0641\u064a \u0627\u0644\u0633\u0627\u0642\u064a\u0646.", "AdobeArabic-Regular", 21, [238, 232, 220], 908, 615, "C", 20],
["Back Bottle", "How to Use EN", "How to Use  |", "MyriadPro-Regular", 15, [43, 30, 14], 818, 732, "L", 0],
["Back Bottle", "How to Use AR", "\u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u0627\u0633\u062a\u062e\u062f\u0627\u0645", "AdobeArabic-Regular", 19, [43, 30, 14], 997, 732, "R", 0],
["Back Bottle", "Directions EN", "Apply a suitable amount on clean,\rdry skin. Gently massage in upward\rcircular motions until fully absorbed.", "MyriadPro-Regular", 14, [238, 232, 220], 907, 758, "C", 15.5],
["Back Bottle", "Directions AR", "\u064a\u0648\u0636\u0639 \u0645\u0642\u062f\u0627\u0631 \u0645\u0646\u0627\u0633\u0628 \u0639\u0644\u0649 \u0628\u0634\u0631\u0629 \u0646\u0638\u064a\u0641\u0629 \u0648\u062c\u0627\u0641\u0629\u060c\r\u0648\u064a\u064f\u062f\u0644\u064e\u0651\u0643 \u0628\u0644\u0637\u0641 \u0628\u062d\u0631\u0643\u0627\u062a \u062f\u0627\u0626\u0631\u064a\u0629 \u0625\u0644\u0649 \u0623\u0639\u0644\u0649 \u062d\u062a\u0649\r\u064a\u062a\u0645 \u0627\u0644\u0627\u0645\u062a\u0635\u0627\u0635 \u0628\u0627\u0644\u0643\u0627\u0645\u0644.", "AdobeArabic-Regular", 17, [238, 232, 220], 909, 808, "C", 14.5],
["Back Bottle", "Ingredients EN", "Ingredients   |", "MyriadPro-Regular", 15, [43, 30, 14], 828, 871, "L", 0],
["Back Bottle", "Ingredients AR", "\u0627\u0644\u0645\u0643\u0648\u0646\u0627\u062a", "AdobeArabic-Regular", 19, [43, 30, 14], 984, 871, "R", 0],
["Back Bottle", "Ingredients list EN", "Cypress Oil, Vitamin E,\rNatural Oils, Botanical Extracts.", "MyriadPro-Regular", 14.5, [238, 232, 220], 810, 893, "L", 15],
["Back Bottle", "Ingredients list AR", "\u0632\u064a\u062a \u0627\u0644\u0633\u0631\u0648\u060c \u0641\u064a\u062a\u0627\u0645\u064a\u0646 E\u060c\r\u0632\u064a\u0648\u062a \u0637\u0628\u064a\u0639\u064a\u0629\u060c \u0648\u0645\u0633\u062a\u062e\u0644\u0635\u0627\u062a \u0646\u0628\u0627\u062a\u064a\u0629.", "AdobeArabic-Regular", 18, [238, 232, 220], 907, 928, "C", 15],
["Back Bottle", "Cautions EN", "Cautions   |", "MyriadPro-Regular", 15, [43, 30, 14], 837, 982, "L", 0],
["Back Bottle", "Cautions AR", "\u062a\u062d\u0630\u064a\u0631\u0627\u062a", "AdobeArabic-Regular", 19, [43, 30, 14], 976, 982, "R", 0],
["Back Bottle", "Cautions list EN", "\u2022  For external use only.\r\u2022  Avoid contact with eyes.\r\u2022  Keep out of reach of children.\r\u2022  Store in a cool, dry place.", "MyriadPro-Regular", 14.5, [238, 232, 220], 812, 1007, "L", 15.7],
["Back Bottle", "Cautions list AR", "\u0644\u0644\u0627\u0633\u062a\u062e\u062f\u0627\u0645 \u0627\u0644\u062e\u0627\u0631\u062c\u064a \u0641\u0642\u0637.\r\u062a\u062c\u0646\u0628 \u0645\u0644\u0627\u0645\u0633\u0629 \u0627\u0644\u0639\u064a\u0646\u064a\u0646.\r\u064a\u062d\u0641\u0638 \u0628\u0639\u064a\u062f\u0627\u064b \u0639\u0646 \u0645\u062a\u0646\u0627\u0648\u0644 \u0627\u0644\u0623\u0637\u0641\u0627\u0644.\r\u064a\u062d\u0641\u0638 \u0641\u064a \u0645\u0643\u0627\u0646 \u0628\u0627\u0631\u062f \u0648\u062c\u0627\u0641.", "AdobeArabic-Regular", 17, [238, 232, 220], 985, 1079, "R", 16],
["Back Bottle", "Volume", "100 ml", "MyriadPro-Regular", 23, [238, 232, 220], 817, 1217, "L", 0]
];

function findSet(name) {
    for (var i = 0; i < doc.layerSets.length; i++) if (doc.layerSets[i].name == name) return doc.layerSets[i];
    var s = doc.layerSets.add(); s.name = name; return s;
}
var textSets = {};
function textSet(group) {
    if (!textSets[group]) {
        var g = findSet(group);
        var t = g.layerSets.add(); t.name = "Text (editable)";
        textSets[group] = t;
    }
    return textSets[group];
}
for (var i = 0; i < S.length; i++) {
    var s = S[i];
    var set = textSet(s[0]);
    var lyr = set.artLayers.add();
    lyr.kind = LayerKind.TEXT;
    var ti = lyr.textItem;
    ti.kind = TextType.POINTTEXT;
    ti.contents = s[2];
    ti.font = s[3];
    ti.size = new UnitValue(s[4], "px");
    var c = new SolidColor(); c.rgb.red = s[5][0]; c.rgb.green = s[5][1]; c.rgb.blue = s[5][2];
    ti.color = c;
    ti.justification = s[8] == "C" ? Justification.CENTER : (s[8] == "R" ? Justification.RIGHT : Justification.LEFT);
    if (s[9] > 0) { ti.useAutoLeading = false; ti.leading = new UnitValue(s[9], "px"); }
    ti.position = [new UnitValue(s[6], "px"), new UnitValue(s[7], "px")];
    lyr.name = s[1];
}
if (res != 72) doc.resizeImage(undefined, undefined, res, ResampleMethod.NONE);
app.preferences.rulerUnits = oldRuler;
app.preferences.typeUnits = oldType;
alert("Done: " + S.length + " editable text layers added.");
