// src/common/relationship.ts
function parseRelationships(root, xml) {
  return xml.elements(root).map((e) => ({
    id: xml.attr(e, "Id"),
    type: xml.attr(e, "Type"),
    target: xml.attr(e, "Target"),
    targetMode: xml.attr(e, "TargetMode")
  }));
}

// src/utils.ts
function escapeClassName(className) {
  return className?.replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase();
}
function encloseFontFamily(fontFamily) {
  return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
}
function splitPath(path) {
  let si = path.lastIndexOf("/") + 1;
  let folder = si == 0 ? "" : path.substring(0, si);
  let fileName = si == 0 ? path : path.substring(si);
  return [folder, fileName];
}
function resolvePath(path, base) {
  try {
    const prefix = "http://docx/";
    const url = new URL(path, prefix + base).toString();
    return url.substring(prefix.length);
  } catch {
    return `${base}${path}`;
  }
}
function keyBy(array, by) {
  return array.reduce((a, x) => {
    a[by(x)] = x;
    return a;
  }, {});
}
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = () => reject();
    reader.readAsDataURL(blob);
  });
}
function isObject(item) {
  return item && typeof item === "object" && !Array.isArray(item);
}
function isString(item) {
  return typeof item === "string" || item instanceof String;
}
function mergeDeep(target, ...sources) {
  if (!sources.length)
    return target;
  const source = sources.shift();
  if (isObject(target) && isObject(source)) {
    for (const key in source) {
      if (isObject(source[key])) {
        const val = target[key] ?? (target[key] = {});
        mergeDeep(val, source[key]);
      } else {
        target[key] = source[key];
      }
    }
  }
  return mergeDeep(target, ...sources);
}
function parseCssRules(text) {
  const result = {};
  for (const rule of text.split(";")) {
    const [key, val] = rule.split(":");
    result[key] = val;
  }
  return result;
}
function formatCssRules(style) {
  return Object.entries(style).map(([k, v], i) => `${k}: ${v}`).join(";");
}
function asArray(val) {
  return Array.isArray(val) ? val : [val];
}
function clamp(val, min, max) {
  return min > val ? min : max < val ? max : val;
}

// src/document/common.ts
var ns = {
  wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
  picture: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  compatibility: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  math: "http://schemas.openxmlformats.org/officeDocument/2006/math"
};
var LengthUsage = {
  Dxa: { mul: 0.05, unit: "pt" },
  //twips
  Emu: { mul: 1 / 12700, unit: "pt" },
  FontSize: { mul: 0.5, unit: "pt" },
  Border: { mul: 0.125, unit: "pt", min: 0.25, max: 12 },
  //NOTE: http://officeopenxml.com/WPtextBorders.php
  Point: { mul: 1, unit: "pt" },
  Percent: { mul: 0.02, unit: "%" },
  LineHeight: { mul: 1 / 240, unit: "" },
  VmlEmu: { mul: 1 / 12700, unit: "" }
};
function convertLength(val, usage = LengthUsage.Dxa) {
  if (val == null || /.+(p[xt]|[%])$/.test(val)) {
    return val;
  }
  var num = parseInt(val) * usage.mul;
  if (usage.min && usage.max)
    num = clamp(num, usage.min, usage.max);
  return `${num.toFixed(2)}${usage.unit}`;
}
function convertBoolean(v, defaultValue = false) {
  switch (v) {
    case "1":
      return true;
    case "0":
      return false;
    case "on":
      return true;
    case "off":
      return false;
    case "true":
      return true;
    case "false":
      return false;
    default:
      return defaultValue;
  }
}
function parseCommonProperty(elem, props, xml) {
  if (elem.namespaceURI != ns.wordml)
    return false;
  switch (elem.localName) {
    case "color":
      props.color = xml.attr(elem, "val");
      break;
    case "sz":
      props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
      break;
    default:
      return false;
  }
  return true;
}

// src/parser/xml-parser.ts
function parseXmlString(xmlString, trimXmlDeclaration = false) {
  if (trimXmlDeclaration)
    xmlString = xmlString.replace(/<[?].*[?]>/, "");
  xmlString = removeUTF8BOM(xmlString);
  const result = new DOMParser().parseFromString(xmlString, "application/xml");
  const errorText = hasXmlParserError(result);
  if (errorText)
    throw new Error(errorText);
  return result;
}
function hasXmlParserError(doc) {
  return doc.getElementsByTagName("parsererror")[0]?.textContent;
}
function removeUTF8BOM(data) {
  return data.charCodeAt(0) === 65279 ? data.substring(1) : data;
}
function serializeXmlString(elem) {
  return new XMLSerializer().serializeToString(elem);
}
var XmlParser = class {
  elements(elem, localName = null) {
    const result = [];
    for (let i = 0, l = elem.childNodes.length; i < l; i++) {
      let c = elem.childNodes.item(i);
      if (c.nodeType == Node.ELEMENT_NODE && (localName == null || c.localName == localName))
        result.push(c);
    }
    return result;
  }
  element(elem, localName) {
    for (let i = 0, l = elem.childNodes.length; i < l; i++) {
      let c = elem.childNodes.item(i);
      if (c.nodeType == 1 && c.localName == localName)
        return c;
    }
    return null;
  }
  elementAttr(elem, localName, attrLocalName) {
    var el = this.element(elem, localName);
    return el ? this.attr(el, attrLocalName) : void 0;
  }
  attrs(elem) {
    return Array.from(elem.attributes);
  }
  attr(elem, localName) {
    for (let i = 0, l = elem.attributes.length; i < l; i++) {
      let a = elem.attributes.item(i);
      if (a.localName == localName)
        return a.value;
    }
    return null;
  }
  intAttr(node, attrName, defaultValue = null) {
    var val = this.attr(node, attrName);
    return val ? parseInt(val) : defaultValue;
  }
  hexAttr(node, attrName, defaultValue = null) {
    var val = this.attr(node, attrName);
    return val ? parseInt(val, 16) : defaultValue;
  }
  floatAttr(node, attrName, defaultValue = null) {
    var val = this.attr(node, attrName);
    return val ? parseFloat(val) : defaultValue;
  }
  boolAttr(node, attrName, defaultValue = null) {
    return convertBoolean(this.attr(node, attrName), defaultValue);
  }
  lengthAttr(node, attrName, usage = LengthUsage.Dxa) {
    return convertLength(this.attr(node, attrName), usage);
  }
};
var globalXmlParser = new XmlParser();
var xml_parser_default = globalXmlParser;

// src/common/part.ts
var Part = class {
  constructor(_package, path) {
    this._package = _package;
    this.path = path;
  }
  async load() {
    this.rels = await this._package.loadRelationships(this.path);
    const xmlText = await this._package.load(this.path);
    const xmlDoc = this._package.parseXmlDocument(xmlText);
    if (this._package.options.keepOrigin) {
      this._xmlDocument = xmlDoc;
    }
    this.parseXml(xmlDoc.firstElementChild);
  }
  save() {
    this._package.update(this.path, serializeXmlString(this._xmlDocument));
  }
  parseXml(root) {
  }
};

// src/font-table/fonts.ts
var embedFontTypeMap = {
  embedRegular: "regular",
  embedBold: "bold",
  embedItalic: "italic",
  embedBoldItalic: "boldItalic"
};
function parseFonts(root, xml) {
  return xml.elements(root).map((el) => parseFont(el, xml));
}
function parseFont(elem, xml) {
  let result = {
    name: xml.attr(elem, "name"),
    embedFontRefs: []
  };
  for (let el of xml.elements(elem)) {
    switch (el.localName) {
      case "family":
        result.family = xml.attr(el, "val");
        break;
      case "altName":
        result.altName = xml.attr(el, "val");
        break;
      case "embedRegular":
      case "embedBold":
      case "embedItalic":
      case "embedBoldItalic":
        result.embedFontRefs.push(parseEmbedFontRef(el, xml));
        break;
    }
  }
  return result;
}
function parseEmbedFontRef(elem, xml) {
  return {
    id: xml.attr(elem, "id"),
    key: xml.attr(elem, "fontKey"),
    type: embedFontTypeMap[elem.localName]
  };
}

// src/font-table/font-table.ts
var FontTablePart = class extends Part {
  parseXml(root) {
    this.fonts = parseFonts(root, this._package.xmlParser);
  }
};

// src/common/open-xml-package.ts
import JSZip from "jszip";
var OpenXmlPackage = class _OpenXmlPackage {
  constructor(_zip, options) {
    this._zip = _zip;
    this.options = options;
    this.xmlParser = new XmlParser();
  }
  get(path) {
    const p = normalizePath(path);
    return this._zip.files[p] ?? this._zip.files[p.replace(/\//g, "\\")];
  }
  update(path, content) {
    this._zip.file(path, content);
  }
  static async load(input, options) {
    const zip = await JSZip.loadAsync(input);
    return new _OpenXmlPackage(zip, options);
  }
  save(type = "blob") {
    return this._zip.generateAsync({ type });
  }
  load(path, type = "string") {
    return this.get(path)?.async(type) ?? Promise.resolve(null);
  }
  async loadRelationships(path = null) {
    let relsPath = `_rels/.rels`;
    if (path != null) {
      const [f, fn] = splitPath(path);
      relsPath = `${f}_rels/${fn}.rels`;
    }
    const txt = await this.load(relsPath);
    return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
  }
  /** @internal */
  parseXmlDocument(txt) {
    return parseXmlString(txt, this.options.trimXmlDeclaration);
  }
};
function normalizePath(path) {
  return path.startsWith("/") ? path.substr(1) : path;
}

// src/document/document-part.ts
var DocumentPart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
  parseXml(root) {
    this.body = this._documentParser.parseDocumentFile(root);
  }
};

// src/document/border.ts
function parseBorder(elem, xml) {
  return {
    type: xml.attr(elem, "val"),
    color: xml.attr(elem, "color"),
    size: xml.lengthAttr(elem, "sz", LengthUsage.Border),
    offset: xml.lengthAttr(elem, "space", LengthUsage.Point),
    frame: xml.boolAttr(elem, "frame"),
    shadow: xml.boolAttr(elem, "shadow")
  };
}
function parseBorders(elem, xml) {
  var result = {};
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "left":
        result.left = parseBorder(e, xml);
        break;
      case "top":
        result.top = parseBorder(e, xml);
        break;
      case "right":
        result.right = parseBorder(e, xml);
        break;
      case "bottom":
        result.bottom = parseBorder(e, xml);
        break;
    }
  }
  return result;
}

// src/document/section.ts
function parseSectionProperties(elem, xml = xml_parser_default) {
  var section = {};
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "pgSz":
        section.pageSize = {
          width: xml.lengthAttr(e, "w"),
          height: xml.lengthAttr(e, "h"),
          orientation: xml.attr(e, "orient")
        };
        break;
      case "type":
        section.type = xml.attr(e, "val");
        break;
      case "pgMar":
        section.pageMargins = {
          left: xml.lengthAttr(e, "left"),
          right: xml.lengthAttr(e, "right"),
          top: xml.lengthAttr(e, "top"),
          bottom: xml.lengthAttr(e, "bottom"),
          header: xml.lengthAttr(e, "header"),
          footer: xml.lengthAttr(e, "footer"),
          gutter: xml.lengthAttr(e, "gutter")
        };
        break;
      case "cols":
        section.columns = parseColumns(e, xml);
        break;
      case "headerReference":
        (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
        break;
      case "footerReference":
        (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
        break;
      case "titlePg":
        section.titlePage = xml.boolAttr(e, "val", true);
        break;
      case "pgBorders":
        section.pageBorders = parseBorders(e, xml);
        break;
      case "pgNumType":
        section.pageNumber = parsePageNumber(e, xml);
        break;
    }
  }
  return section;
}
function parseColumns(elem, xml) {
  return {
    numberOfColumns: xml.intAttr(elem, "num"),
    space: xml.lengthAttr(elem, "space"),
    separator: xml.boolAttr(elem, "sep"),
    equalWidth: xml.boolAttr(elem, "equalWidth", true),
    columns: xml.elements(elem, "col").map((e) => ({
      width: xml.lengthAttr(e, "w"),
      space: xml.lengthAttr(e, "space")
    }))
  };
}
function parsePageNumber(elem, xml) {
  return {
    chapSep: xml.attr(elem, "chapSep"),
    chapStyle: xml.attr(elem, "chapStyle"),
    format: xml.attr(elem, "fmt"),
    start: xml.intAttr(elem, "start")
  };
}
function parseFooterHeaderReference(elem, xml) {
  return {
    id: xml.attr(elem, "id"),
    type: xml.attr(elem, "type")
  };
}

// src/document/line-spacing.ts
function parseLineSpacing(elem, xml) {
  return {
    before: xml.lengthAttr(elem, "before"),
    after: xml.lengthAttr(elem, "after"),
    line: xml.intAttr(elem, "line"),
    lineRule: xml.attr(elem, "lineRule")
  };
}

// src/document/run.ts
function parseRunProperties(elem, xml) {
  let result = {};
  for (let el of xml.elements(elem)) {
    parseRunProperty(el, result, xml);
  }
  return result;
}
function parseRunProperty(elem, props, xml) {
  if (parseCommonProperty(elem, props, xml))
    return true;
  return false;
}

// src/document/paragraph.ts
function parseParagraphProperties(elem, xml) {
  let result = {};
  for (let el of xml.elements(elem)) {
    parseParagraphProperty(el, result, xml);
  }
  return result;
}
function parseParagraphProperty(elem, props, xml) {
  if (elem.namespaceURI != ns.wordml)
    return false;
  if (parseCommonProperty(elem, props, xml))
    return true;
  switch (elem.localName) {
    case "tabs":
      props.tabs = parseTabs(elem, xml);
      break;
    case "sectPr":
      props.sectionProps = parseSectionProperties(elem, xml);
      break;
    case "numPr":
      props.numbering = parseNumbering(elem, xml);
      break;
    case "spacing":
      props.lineSpacing = parseLineSpacing(elem, xml);
      return false;
      break;
    case "textAlignment":
      props.textAlignment = xml.attr(elem, "val");
      return false;
      break;
    case "keepLines":
      props.keepLines = xml.boolAttr(elem, "val", true);
      break;
    case "keepNext":
      props.keepNext = xml.boolAttr(elem, "val", true);
      break;
    case "pageBreakBefore":
      props.pageBreakBefore = xml.boolAttr(elem, "val", true);
      break;
    case "outlineLvl":
      props.outlineLevel = xml.intAttr(elem, "val");
      break;
    case "pStyle":
      props.styleName = xml.attr(elem, "val");
      break;
    case "rPr":
      props.runProps = parseRunProperties(elem, xml);
      break;
    default:
      return false;
  }
  return true;
}
function parseTabs(elem, xml) {
  return xml.elements(elem, "tab").map((e) => ({
    position: xml.lengthAttr(e, "pos"),
    leader: xml.attr(e, "leader"),
    style: xml.attr(e, "val")
  }));
}
function parseNumbering(elem, xml) {
  var result = {};
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "numId":
        result.id = xml.attr(e, "val");
        break;
      case "ilvl":
        result.level = xml.intAttr(e, "val");
        break;
    }
  }
  return result;
}

// src/numbering/numbering.ts
function parseNumberingPart(elem, xml) {
  let result = {
    numberings: [],
    abstractNumberings: [],
    bulletPictures: []
  };
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "num":
        result.numberings.push(parseNumbering2(e, xml));
        break;
      case "abstractNum":
        result.abstractNumberings.push(parseAbstractNumbering(e, xml));
        break;
      case "numPicBullet":
        result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
        break;
    }
  }
  return result;
}
function parseNumbering2(elem, xml) {
  let result = {
    id: xml.attr(elem, "numId"),
    overrides: []
  };
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "abstractNumId":
        result.abstractId = xml.attr(e, "val");
        break;
      case "lvlOverride":
        result.overrides.push(parseNumberingLevelOverrride(e, xml));
        break;
    }
  }
  return result;
}
function parseAbstractNumbering(elem, xml) {
  let result = {
    id: xml.attr(elem, "abstractNumId"),
    levels: []
  };
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "name":
        result.name = xml.attr(e, "val");
        break;
      case "multiLevelType":
        result.multiLevelType = xml.attr(e, "val");
        break;
      case "numStyleLink":
        result.numberingStyleLink = xml.attr(e, "val");
        break;
      case "styleLink":
        result.styleLink = xml.attr(e, "val");
        break;
      case "lvl":
        result.levels.push(parseNumberingLevel(e, xml));
        break;
    }
  }
  return result;
}
function parseNumberingLevel(elem, xml) {
  let result = {
    level: xml.intAttr(elem, "ilvl")
  };
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "start":
        result.start = xml.attr(e, "val");
        break;
      case "lvlRestart":
        result.restart = xml.intAttr(e, "val");
        break;
      case "numFmt":
        result.format = xml.attr(e, "val");
        break;
      case "lvlText":
        result.text = xml.attr(e, "val");
        break;
      case "lvlJc":
        result.justification = xml.attr(e, "val");
        break;
      case "lvlPicBulletId":
        result.bulletPictureId = xml.attr(e, "val");
        break;
      case "pStyle":
        result.paragraphStyle = xml.attr(e, "val");
        break;
      case "pPr":
        result.paragraphProps = parseParagraphProperties(e, xml);
        break;
      case "rPr":
        result.runProps = parseRunProperties(e, xml);
        break;
    }
  }
  return result;
}
function parseNumberingLevelOverrride(elem, xml) {
  let result = {
    level: xml.intAttr(elem, "ilvl")
  };
  for (let e of xml.elements(elem)) {
    switch (e.localName) {
      case "startOverride":
        result.start = xml.intAttr(e, "val");
        break;
      case "lvl":
        result.numberingLevel = parseNumberingLevel(e, xml);
        break;
    }
  }
  return result;
}
function parseNumberingBulletPicture(elem, xml) {
  var pict = xml.element(elem, "pict");
  var shape = pict && xml.element(pict, "shape");
  var imagedata = shape && xml.element(shape, "imagedata");
  return imagedata ? {
    id: xml.attr(elem, "numPicBulletId"),
    referenceId: xml.attr(imagedata, "id"),
    style: xml.attr(shape, "style")
  } : null;
}

// src/numbering/numbering-part.ts
var NumberingPart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
  parseXml(root) {
    Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
    this.domNumberings = this._documentParser.parseNumberingFile(root);
  }
};

// src/styles/styles-part.ts
var StylesPart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
  parseXml(root) {
    this.styles = this._documentParser.parseStylesFile(root);
  }
};

// src/document/dom.ts
var OpenXmlElementBase = class {
  constructor() {
    this.children = [];
    this.cssStyle = {};
  }
};

// src/header-footer/elements.ts
var WmlHeader = class extends OpenXmlElementBase {
  constructor() {
    super(...arguments);
    this.type = "header" /* Header */;
  }
};
var WmlFooter = class extends OpenXmlElementBase {
  constructor() {
    super(...arguments);
    this.type = "footer" /* Footer */;
  }
};

// src/header-footer/parts.ts
var BaseHeaderFooterPart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
  parseXml(root) {
    this.rootElement = this.createRootElement();
    this.rootElement.children = this._documentParser.parseBodyElements(root);
  }
};
var HeaderPart = class extends BaseHeaderFooterPart {
  createRootElement() {
    return new WmlHeader();
  }
};
var FooterPart = class extends BaseHeaderFooterPart {
  createRootElement() {
    return new WmlFooter();
  }
};

// src/document-props/extended-props.ts
function parseExtendedProps(root, xmlParser) {
  const result = {};
  for (let el of xmlParser.elements(root)) {
    switch (el.localName) {
      case "Template":
        result.template = el.textContent;
        break;
      case "Pages":
        result.pages = safeParseToInt(el.textContent);
        break;
      case "Words":
        result.words = safeParseToInt(el.textContent);
        break;
      case "Characters":
        result.characters = safeParseToInt(el.textContent);
        break;
      case "Application":
        result.application = el.textContent;
        break;
      case "Lines":
        result.lines = safeParseToInt(el.textContent);
        break;
      case "Paragraphs":
        result.paragraphs = safeParseToInt(el.textContent);
        break;
      case "Company":
        result.company = el.textContent;
        break;
      case "AppVersion":
        result.appVersion = el.textContent;
        break;
    }
  }
  return result;
}
function safeParseToInt(value) {
  if (typeof value === "undefined")
    return;
  return parseInt(value);
}

// src/document-props/extended-props-part.ts
var ExtendedPropsPart = class extends Part {
  parseXml(root) {
    this.props = parseExtendedProps(root, this._package.xmlParser);
  }
};

// src/document-props/core-props.ts
function parseCoreProps(root, xmlParser) {
  const result = {};
  for (let el of xmlParser.elements(root)) {
    switch (el.localName) {
      case "title":
        result.title = el.textContent;
        break;
      case "description":
        result.description = el.textContent;
        break;
      case "subject":
        result.subject = el.textContent;
        break;
      case "creator":
        result.creator = el.textContent;
        break;
      case "keywords":
        result.keywords = el.textContent;
        break;
      case "language":
        result.language = el.textContent;
        break;
      case "lastModifiedBy":
        result.lastModifiedBy = el.textContent;
        break;
      case "revision":
        el.textContent && (result.revision = parseInt(el.textContent));
        break;
    }
  }
  return result;
}

// src/document-props/core-props-part.ts
var CorePropsPart = class extends Part {
  parseXml(root) {
    this.props = parseCoreProps(root, this._package.xmlParser);
  }
};

// src/theme/theme.ts
var DmlTheme = class {
};
function parseTheme(elem, xml) {
  var result = new DmlTheme();
  var themeElements = xml.element(elem, "themeElements");
  for (let el of xml.elements(themeElements)) {
    switch (el.localName) {
      case "clrScheme":
        result.colorScheme = parseColorScheme(el, xml);
        break;
      case "fontScheme":
        result.fontScheme = parseFontScheme(el, xml);
        break;
    }
  }
  return result;
}
function parseColorScheme(elem, xml) {
  var result = {
    name: xml.attr(elem, "name"),
    colors: {}
  };
  for (let el of xml.elements(elem)) {
    var srgbClr = xml.element(el, "srgbClr");
    var sysClr = xml.element(el, "sysClr");
    if (srgbClr) {
      result.colors[el.localName] = xml.attr(srgbClr, "val");
    } else if (sysClr) {
      result.colors[el.localName] = xml.attr(sysClr, "lastClr");
    }
  }
  return result;
}
function parseFontScheme(elem, xml) {
  var result = {
    name: xml.attr(elem, "name")
  };
  for (let el of xml.elements(elem)) {
    switch (el.localName) {
      case "majorFont":
        result.majorFont = parseFontInfo(el, xml);
        break;
      case "minorFont":
        result.minorFont = parseFontInfo(el, xml);
        break;
    }
  }
  return result;
}
function parseFontInfo(elem, xml) {
  return {
    latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
    eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
    csTypeface: xml.elementAttr(elem, "cs", "typeface")
  };
}

// src/theme/theme-part.ts
var ThemePart = class extends Part {
  constructor(pkg, path) {
    super(pkg, path);
  }
  parseXml(root) {
    this.theme = parseTheme(root, this._package.xmlParser);
  }
};

// src/notes/elements.ts
var WmlBaseNote = class {
};
var WmlFootnote = class extends WmlBaseNote {
  constructor() {
    super(...arguments);
    this.type = "footnote" /* Footnote */;
  }
};
var WmlEndnote = class extends WmlBaseNote {
  constructor() {
    super(...arguments);
    this.type = "endnote" /* Endnote */;
  }
};

// src/notes/parts.ts
var BaseNotePart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
};
var FootnotesPart = class extends BaseNotePart {
  constructor(pkg, path, parser) {
    super(pkg, path, parser);
  }
  parseXml(root) {
    this.notes = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
  }
};
var EndnotesPart = class extends BaseNotePart {
  constructor(pkg, path, parser) {
    super(pkg, path, parser);
  }
  parseXml(root) {
    this.notes = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
  }
};

// src/settings/settings.ts
function parseSettings(elem, xml) {
  var result = {};
  for (let el of xml.elements(elem)) {
    switch (el.localName) {
      case "defaultTabStop":
        result.defaultTabStop = xml.lengthAttr(el, "val");
        break;
      case "footnotePr":
        result.footnoteProps = parseNoteProperties(el, xml);
        break;
      case "endnotePr":
        result.endnoteProps = parseNoteProperties(el, xml);
        break;
      case "autoHyphenation":
        result.autoHyphenation = xml.boolAttr(el, "val");
        break;
    }
  }
  return result;
}
function parseNoteProperties(elem, xml) {
  var result = {
    defaultNoteIds: []
  };
  for (let el of xml.elements(elem)) {
    switch (el.localName) {
      case "numFmt":
        result.nummeringFormat = xml.attr(el, "val");
        break;
      case "footnote":
      case "endnote":
        result.defaultNoteIds.push(xml.attr(el, "id"));
        break;
    }
  }
  return result;
}

// src/settings/settings-part.ts
var SettingsPart = class extends Part {
  constructor(pkg, path) {
    super(pkg, path);
  }
  parseXml(root) {
    this.settings = parseSettings(root, this._package.xmlParser);
  }
};

// src/document-props/custom-props.ts
function parseCustomProps(root, xml) {
  return xml.elements(root, "property").map((e) => {
    const firstChild = e.firstChild;
    return {
      formatId: xml.attr(e, "fmtid"),
      name: xml.attr(e, "name"),
      type: firstChild.nodeName,
      value: firstChild.textContent
    };
  });
}

// src/document-props/custom-props-part.ts
var CustomPropsPart = class extends Part {
  parseXml(root) {
    this.props = parseCustomProps(root, this._package.xmlParser);
  }
};

// src/comments/comments-part.ts
var CommentsPart = class extends Part {
  constructor(pkg, path, parser) {
    super(pkg, path);
    this._documentParser = parser;
  }
  parseXml(root) {
    this.comments = this._documentParser.parseComments(root);
    this.commentMap = keyBy(this.comments, (x) => x.id);
  }
};

// src/comments/comments-extended-part.ts
var CommentsExtendedPart = class extends Part {
  constructor(pkg, path) {
    super(pkg, path);
    this.comments = [];
  }
  parseXml(root) {
    const xml = this._package.xmlParser;
    for (let el of xml.elements(root, "commentEx")) {
      this.comments.push({
        paraId: xml.attr(el, "paraId"),
        paraIdParent: xml.attr(el, "paraIdParent"),
        done: xml.boolAttr(el, "done")
      });
    }
    this.commentMap = keyBy(this.comments, (x) => x.paraId);
  }
};

// src/word-document.ts
var topLevelRels = [
  { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" /* OfficeDocument */, target: "word/document.xml" },
  { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" /* ExtendedProperties */, target: "docProps/app.xml" },
  { type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" /* CoreProperties */, target: "docProps/core.xml" },
  { type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties" /* CustomProperties */, target: "docProps/custom.xml" }
];
var WordDocument = class _WordDocument {
  constructor() {
    this.parts = [];
    this.partsMap = {};
  }
  static async load(blob, parser, options) {
    var d = new _WordDocument();
    d._options = options;
    d._parser = parser;
    d._package = await OpenXmlPackage.load(blob, options);
    d.rels = await d._package.loadRelationships();
    await Promise.all(topLevelRels.map((rel) => {
      const r = d.rels.find((x) => x.type === rel.type) ?? rel;
      return d.loadRelationshipPart(r.target, r.type);
    }));
    return d;
  }
  save(type = "blob") {
    return this._package.save(type);
  }
  async loadRelationshipPart(path, type) {
    if (this.partsMap[path])
      return this.partsMap[path];
    if (!this._package.get(path))
      return null;
    let part = null;
    switch (type) {
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" /* OfficeDocument */:
        this.documentPart = part = new DocumentPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" /* FontTable */:
        this.fontTablePart = part = new FontTablePart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" /* Numbering */:
        this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" /* Styles */:
        this.stylesPart = part = new StylesPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" /* Theme */:
        this.themePart = part = new ThemePart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" /* Footnotes */:
        this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" /* Endnotes */:
        this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" /* Footer */:
        part = new FooterPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" /* Header */:
        part = new HeaderPart(this._package, path, this._parser);
        break;
      case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" /* CoreProperties */:
        this.corePropsPart = part = new CorePropsPart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" /* ExtendedProperties */:
        this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties" /* CustomProperties */:
        part = new CustomPropsPart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" /* Settings */:
        this.settingsPart = part = new SettingsPart(this._package, path);
        break;
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" /* Comments */:
        this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
        break;
      case "http://schemas.microsoft.com/office/2011/relationships/commentsExtended" /* CommentsExtended */:
        this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
        break;
    }
    if (part == null)
      return Promise.resolve(null);
    this.partsMap[path] = part;
    this.parts.push(part);
    await part.load();
    if (part.rels?.length > 0) {
      const [folder] = splitPath(part.path);
      await Promise.all(part.rels.map((rel) => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
    }
    return part;
  }
  async loadDocumentImage(id, part) {
    const x = await this.loadResource(part ?? this.documentPart, id, "blob");
    return this.blobToURL(x);
  }
  async loadNumberingImage(id) {
    const x = await this.loadResource(this.numberingPart, id, "blob");
    return this.blobToURL(x);
  }
  async loadFont(id, key) {
    const x = await this.loadResource(this.fontTablePart, id, "uint8array");
    return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
  }
  async loadAltChunk(id, part) {
    return await this.loadResource(part ?? this.documentPart, id, "string");
  }
  blobToURL(blob) {
    if (!blob)
      return null;
    if (this._options.useBase64URL) {
      return blobToBase64(blob);
    }
    return URL.createObjectURL(blob);
  }
  findPartByRelId(id, basePart = null) {
    var rel = (basePart.rels ?? this.rels).find((r) => r.id == id);
    const folder = basePart ? splitPath(basePart.path)[0] : "";
    return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
  }
  getPathById(part, id) {
    const rel = part.rels.find((x) => x.id == id);
    const [folder] = splitPath(part.path);
    return rel ? resolvePath(rel.target, folder) : null;
  }
  loadResource(part, id, outputType) {
    const path = this.getPathById(part, id);
    return path ? this._package.load(path, outputType) : Promise.resolve(null);
  }
};
function deobfuscate(data, guidKey) {
  const len = 16;
  const trimmed = guidKey.replace(/{|}|-/g, "");
  const numbers = new Array(len);
  for (let i = 0; i < len; i++)
    numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);
  for (let i = 0; i < 32; i++)
    data[i] = data[i] ^ numbers[i % len];
  return data;
}

// src/document/bookmarks.ts
function parseBookmarkStart(elem, xml) {
  return {
    type: "bookmarkStart" /* BookmarkStart */,
    id: xml.attr(elem, "id"),
    name: xml.attr(elem, "name"),
    colFirst: xml.intAttr(elem, "colFirst"),
    colLast: xml.intAttr(elem, "colLast")
  };
}
function parseBookmarkEnd(elem, xml) {
  return {
    type: "bookmarkEnd" /* BookmarkEnd */,
    id: xml.attr(elem, "id")
  };
}

// src/vml/vml.ts
var VmlElement = class extends OpenXmlElementBase {
  constructor() {
    super(...arguments);
    this.type = "vmlElement" /* VmlElement */;
    this.attrs = {};
  }
};
function parseVmlElement(elem, parser) {
  var result = new VmlElement();
  switch (elem.localName) {
    case "rect":
      result.tagName = "rect";
      Object.assign(result.attrs, { width: "100%", height: "100%" });
      break;
    case "oval":
      result.tagName = "ellipse";
      Object.assign(result.attrs, { cx: "50%", cy: "50%", rx: "50%", ry: "50%" });
      break;
    case "line":
      result.tagName = "line";
      break;
    case "shape":
      result.tagName = "g";
      break;
    case "textbox":
      result.tagName = "foreignObject";
      Object.assign(result.attrs, { width: "100%", height: "100%" });
      result.attrs.class = (result.attrs.class ? result.attrs.class + " " : "") + "textbox";
      break;
    default:
      return null;
  }
  for (const at of xml_parser_default.attrs(elem)) {
    switch (at.localName) {
      case "style":
        result.cssStyle = parseCssRules(at.value || "");
        if (result.cssStyle["mso-position-horizontal-relative"] === "page") {
          result.cssStyle["position"] = "absolute";
          result.cssStyle["left"] = result.cssStyle["mso-position-horizontal"] || "0";
          delete result.cssStyle["mso-position-horizontal-relative"];
          delete result.cssStyle["mso-position-horizontal"];
        }
        if (result.cssStyle["mso-position-vertical-relative"] === "page") {
          result.cssStyle["position"] = "absolute";
          result.cssStyle["top"] = result.cssStyle["mso-position-vertical"] || "0";
          delete result.cssStyle["mso-position-vertical-relative"];
          delete result.cssStyle["mso-position-vertical"];
        }
        if (result.cssStyle["left"] === "absolute") {
          result.cssStyle["left"] = "0";
        }
        if (result.cssStyle["top"] === "absolute") {
          result.cssStyle["top"] = "0";
        }
        if (result.cssStyle["right"] === "absolute") {
          result.cssStyle["right"] = "0";
        }
        if (result.cssStyle["bottom"] === "absolute") {
          result.cssStyle["bottom"] = "0";
        }
        break;
      case "fillcolor":
        result.attrs.fill = at.value;
        break;
      case "from":
        const [x1, y1] = parsePoint(at.value);
        Object.assign(result.attrs, { x1, y1 });
        break;
      case "to":
        const [x2, y2] = parsePoint(at.value);
        Object.assign(result.attrs, { x2, y2 });
        break;
    }
  }
  for (const el of xml_parser_default.elements(elem)) {
    switch (el.localName) {
      case "stroke":
        Object.assign(result.attrs, parseStroke(el));
        break;
      case "fill":
        Object.assign(result.attrs, parseFill(el));
        break;
      case "imagedata":
        result.tagName = "image";
        Object.assign(result.attrs, { width: "100%", height: "100%" });
        result.imageHref = {
          id: xml_parser_default.attr(el, "id"),
          title: xml_parser_default.attr(el, "title")
        };
        break;
      case "txbxContent":
        result.children.push(...parser.parseBodyElements(el));
        break;
      default:
        const child = parseVmlElement(el, parser);
        child && result.children.push(child);
        break;
    }
  }
  return result;
}
function parseStroke(el) {
  return {
    "stroke": xml_parser_default.attr(el, "color"),
    "stroke-width": xml_parser_default.lengthAttr(el, "weight", LengthUsage.Emu) ?? "1px"
  };
}
function parseFill(el) {
  return {
    //'fill': xml.attr(el, "color2")
  };
}
function parsePoint(val) {
  return val.split(",");
}

// src/comments/elements.ts
var WmlComment = class extends OpenXmlElementBase {
  constructor() {
    super(...arguments);
    this.type = "comment" /* Comment */;
  }
};
var WmlCommentReference = class extends OpenXmlElementBase {
  constructor(id) {
    super();
    this.id = id;
    this.type = "commentReference" /* CommentReference */;
  }
};
var WmlCommentRangeStart = class extends OpenXmlElementBase {
  constructor(id) {
    super();
    this.id = id;
    this.type = "commentRangeStart" /* CommentRangeStart */;
  }
};
var WmlCommentRangeEnd = class extends OpenXmlElementBase {
  constructor(id) {
    super();
    this.id = id;
    this.type = "commentRangeEnd" /* CommentRangeEnd */;
  }
};

// src/document-parser.ts
var autos = {
  shd: "inherit",
  color: "black",
  borderColor: "black",
  highlight: "transparent"
};
var supportedNamespaceURIs = [];
var mmlTagMap = {
  "oMath": "mmlMath" /* MmlMath */,
  "oMathPara": "mmlMathParagraph" /* MmlMathParagraph */,
  "f": "mmlFraction" /* MmlFraction */,
  "func": "mmlFunction" /* MmlFunction */,
  "fName": "mmlFunctionName" /* MmlFunctionName */,
  "num": "mmlNumerator" /* MmlNumerator */,
  "den": "mmlDenominator" /* MmlDenominator */,
  "rad": "mmlRadical" /* MmlRadical */,
  "deg": "mmlDegree" /* MmlDegree */,
  "e": "mmlBase" /* MmlBase */,
  "sSup": "mmlSuperscript" /* MmlSuperscript */,
  "sSub": "mmlSubscript" /* MmlSubscript */,
  "sPre": "mmlPreSubSuper" /* MmlPreSubSuper */,
  "sup": "mmlSuperArgument" /* MmlSuperArgument */,
  "sub": "mmlSubArgument" /* MmlSubArgument */,
  "d": "mmlDelimiter" /* MmlDelimiter */,
  "nary": "mmlNary" /* MmlNary */,
  "eqArr": "mmlEquationArray" /* MmlEquationArray */,
  "lim": "mmlLimit" /* MmlLimit */,
  "limLow": "mmlLimitLower" /* MmlLimitLower */,
  "m": "mmlMatrix" /* MmlMatrix */,
  "mr": "mmlMatrixRow" /* MmlMatrixRow */,
  "box": "mmlBox" /* MmlBox */,
  "bar": "mmlBar" /* MmlBar */,
  "groupChr": "mmlGroupChar" /* MmlGroupChar */
};
var DocumentParser = class {
  constructor(options) {
    this.options = {
      ignoreWidth: false,
      debug: false,
      ...options
    };
  }
  parseNotes(xmlDoc, elemName, elemClass) {
    var result = [];
    for (let el of xml_parser_default.elements(xmlDoc, elemName)) {
      const node = new elemClass();
      node.id = xml_parser_default.attr(el, "id");
      node.noteType = xml_parser_default.attr(el, "type");
      node.children = this.parseBodyElements(el);
      result.push(node);
    }
    return result;
  }
  parseComments(xmlDoc) {
    var result = [];
    for (let el of xml_parser_default.elements(xmlDoc, "comment")) {
      const item = new WmlComment();
      item.id = xml_parser_default.attr(el, "id");
      item.author = xml_parser_default.attr(el, "author");
      item.initials = xml_parser_default.attr(el, "initials");
      item.date = xml_parser_default.attr(el, "date");
      item.children = this.parseBodyElements(el);
      result.push(item);
    }
    return result;
  }
  parseDocumentFile(xmlDoc) {
    var xbody = xml_parser_default.element(xmlDoc, "body");
    var background = xml_parser_default.element(xmlDoc, "background");
    var sectPr = xml_parser_default.element(xbody, "sectPr");
    return {
      type: "document" /* Document */,
      children: this.parseBodyElements(xbody),
      props: sectPr ? parseSectionProperties(sectPr, xml_parser_default) : {},
      cssStyle: background ? this.parseBackground(background) : {}
    };
  }
  parseBackground(elem) {
    var result = {};
    var color = xmlUtil.colorAttr(elem, "color");
    if (color) {
      result["background-color"] = color;
    }
    return result;
  }
  parseBodyElements(element) {
    var children = [];
    for (const elem of xml_parser_default.elements(element)) {
      switch (elem.localName) {
        case "p":
          children.push(this.parseParagraph(elem));
          break;
        case "altChunk":
          children.push(this.parseAltChunk(elem));
          break;
        case "tbl":
          children.push(this.parseTable(elem));
          break;
        case "sdt":
          children.push(...this.parseSdt(elem, (e) => this.parseBodyElements(e)));
          break;
      }
    }
    return children;
  }
  parseStylesFile(xstyles) {
    var result = [];
    for (const n of xml_parser_default.elements(xstyles)) {
      switch (n.localName) {
        case "style":
          result.push(this.parseStyle(n));
          break;
        case "docDefaults":
          result.push(this.parseDefaultStyles(n));
          break;
      }
    }
    return result;
  }
  parseDefaultStyles(node) {
    var result = {
      id: null,
      name: null,
      target: null,
      basedOn: null,
      styles: []
    };
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "rPrDefault":
          var rPr = xml_parser_default.element(c, "rPr");
          if (rPr)
            result.styles.push({
              target: "span",
              values: this.parseDefaultProperties(rPr, {})
            });
          break;
        case "pPrDefault":
          var pPr = xml_parser_default.element(c, "pPr");
          if (pPr)
            result.styles.push({
              target: "p",
              values: this.parseDefaultProperties(pPr, {})
            });
          break;
      }
    }
    return result;
  }
  parseStyle(node) {
    var result = {
      id: xml_parser_default.attr(node, "styleId"),
      isDefault: xml_parser_default.boolAttr(node, "default"),
      name: null,
      target: null,
      basedOn: null,
      styles: [],
      linked: null
    };
    switch (xml_parser_default.attr(node, "type")) {
      case "paragraph":
        result.target = "p";
        break;
      case "table":
        result.target = "table";
        break;
      case "character":
        result.target = "span";
        break;
    }
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "basedOn":
          result.basedOn = xml_parser_default.attr(n, "val");
          break;
        case "name":
          result.name = xml_parser_default.attr(n, "val");
          break;
        case "link":
          result.linked = xml_parser_default.attr(n, "val");
          break;
        case "next":
          result.next = xml_parser_default.attr(n, "val");
          break;
        case "aliases":
          result.aliases = xml_parser_default.attr(n, "val").split(",");
          break;
        case "pPr":
          result.styles.push({
            target: "p",
            values: this.parseDefaultProperties(n, {})
          });
          result.paragraphProps = parseParagraphProperties(n, xml_parser_default);
          break;
        case "rPr":
          result.styles.push({
            target: "span",
            values: this.parseDefaultProperties(n, {})
          });
          result.runProps = parseRunProperties(n, xml_parser_default);
          break;
        case "tblPr":
        case "tcPr":
          result.styles.push({
            target: "td",
            //TODO: maybe move to processor
            values: this.parseDefaultProperties(n, {})
          });
          break;
        case "tblStylePr":
          for (let s of this.parseTableStyle(n))
            result.styles.push(s);
          break;
        case "rsid":
        case "qFormat":
        case "hidden":
        case "semiHidden":
        case "unhideWhenUsed":
        case "autoRedefine":
        case "uiPriority":
          break;
        default:
          this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
      }
    }
    return result;
  }
  parseTableStyle(node) {
    var result = [];
    var type = xml_parser_default.attr(node, "type");
    var selector = "";
    var modificator = "";
    switch (type) {
      case "firstRow":
        modificator = ".first-row";
        selector = "tr.first-row td";
        break;
      case "lastRow":
        modificator = ".last-row";
        selector = "tr.last-row td";
        break;
      case "firstCol":
        modificator = ".first-col";
        selector = "td.first-col";
        break;
      case "lastCol":
        modificator = ".last-col";
        selector = "td.last-col";
        break;
      case "band1Vert":
        modificator = ":not(.no-vband)";
        selector = "td.odd-col";
        break;
      case "band2Vert":
        modificator = ":not(.no-vband)";
        selector = "td.even-col";
        break;
      case "band1Horz":
        modificator = ":not(.no-hband)";
        selector = "tr.odd-row";
        break;
      case "band2Horz":
        modificator = ":not(.no-hband)";
        selector = "tr.even-row";
        break;
      default:
        return [];
    }
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "pPr":
          result.push({
            target: `${selector} p`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {})
          });
          break;
        case "rPr":
          result.push({
            target: `${selector} span`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {})
          });
          break;
        case "tblPr":
        case "tcPr":
          result.push({
            target: selector,
            //TODO: maybe move to processor
            mod: modificator,
            values: this.parseDefaultProperties(n, {})
          });
          break;
      }
    }
    return result;
  }
  parseNumberingFile(node) {
    var result = [];
    var mapping = {};
    var bullets = [];
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "abstractNum":
          this.parseAbstractNumbering(n, bullets).forEach((x) => result.push(x));
          break;
        case "numPicBullet":
          bullets.push(this.parseNumberingPicBullet(n));
          break;
        case "num":
          var numId = xml_parser_default.attr(n, "numId");
          var abstractNumId = xml_parser_default.elementAttr(n, "abstractNumId", "val");
          mapping[abstractNumId] = numId;
          break;
      }
    }
    result.forEach((x) => x.id = mapping[x.id]);
    return result;
  }
  parseNumberingPicBullet(elem) {
    var pict = xml_parser_default.element(elem, "pict");
    var shape = pict && xml_parser_default.element(pict, "shape");
    var imagedata = shape && xml_parser_default.element(shape, "imagedata");
    return imagedata ? {
      id: xml_parser_default.intAttr(elem, "numPicBulletId"),
      src: xml_parser_default.attr(imagedata, "id"),
      style: xml_parser_default.attr(shape, "style")
    } : null;
  }
  parseAbstractNumbering(node, bullets) {
    var result = [];
    var id = xml_parser_default.attr(node, "abstractNumId");
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "lvl":
          result.push(this.parseNumberingLevel(id, n, bullets));
          break;
      }
    }
    return result;
  }
  parseNumberingLevel(id, node, bullets) {
    var result = {
      id,
      level: xml_parser_default.intAttr(node, "ilvl"),
      start: 1,
      pStyleName: void 0,
      pStyle: {},
      rStyle: {},
      suff: "tab"
    };
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "start":
          result.start = xml_parser_default.intAttr(n, "val");
          break;
        case "pPr":
          this.parseDefaultProperties(n, result.pStyle);
          break;
        case "rPr":
          this.parseDefaultProperties(n, result.rStyle);
          break;
        case "lvlPicBulletId":
          var bulletId = xml_parser_default.intAttr(n, "val");
          result.bullet = bullets.find((x) => x?.id == bulletId);
          break;
        case "lvlText":
          result.levelText = xml_parser_default.attr(n, "val");
          break;
        case "pStyle":
          result.pStyleName = xml_parser_default.attr(n, "val");
          break;
        case "numFmt":
          result.format = xml_parser_default.attr(n, "val");
          break;
        case "suff":
          result.suff = xml_parser_default.attr(n, "val");
          break;
      }
    }
    if (result.pStyle["margin-inline-start"] && result.pStyle["text-indent"].startsWith("-")) {
      const value = result.pStyle["text-indent"].substring(1);
      result.rStyle["display"] = "inline-block";
      result.rStyle["width"] = value;
      result.rStyle["text-align"] = "end";
      result.rStyle["padding-inline-end"] = `calc(${value} / 2)`;
    }
    return result;
  }
  parseSdt(node, parser) {
    const sdtContent = xml_parser_default.element(node, "sdtContent");
    return sdtContent ? parser(sdtContent) : [];
  }
  parseInserted(node, parentParser) {
    return {
      type: "inserted" /* Inserted */,
      children: parentParser(node)?.children ?? []
    };
  }
  parseDeleted(node, parentParser) {
    return {
      type: "deleted" /* Deleted */,
      children: parentParser(node)?.children ?? []
    };
  }
  parseAltChunk(node) {
    return { type: "altChunk" /* AltChunk */, children: [], id: xml_parser_default.attr(node, "id") };
  }
  parseParagraph(node) {
    var result = { type: "paragraph" /* Paragraph */, children: [] };
    for (let el of xml_parser_default.elements(node)) {
      switch (el.localName) {
        case "pPr":
          this.parseParagraphProperties(el, result);
          break;
        case "r":
          result.children.push(this.parseRun(el, result));
          break;
        case "hyperlink":
          result.children.push(this.parseHyperlink(el, result));
          break;
        case "smartTag":
          result.children.push(this.parseSmartTag(el, result));
          break;
        case "bookmarkStart":
          result.children.push(parseBookmarkStart(el, xml_parser_default));
          break;
        case "bookmarkEnd":
          result.children.push(parseBookmarkEnd(el, xml_parser_default));
          break;
        case "commentRangeStart":
          result.children.push(new WmlCommentRangeStart(xml_parser_default.attr(el, "id")));
          break;
        case "commentRangeEnd":
          result.children.push(new WmlCommentRangeEnd(xml_parser_default.attr(el, "id")));
          break;
        case "oMath":
        case "oMathPara":
          result.children.push(this.parseMathElement(el));
          break;
        case "sdt":
          result.children.push(...this.parseSdt(el, (e) => this.parseParagraph(e).children));
          break;
        case "ins":
          result.children.push(this.parseInserted(el, (e) => this.parseParagraph(e)));
          break;
        case "del":
          result.children.push(this.parseDeleted(el, (e) => this.parseParagraph(e)));
          break;
      }
    }
    if (result.children.length == 0) {
      if (result.cssStyle && result.cssStyle["text-decoration"]) result.cssStyle["text-decoration"] = "none";
      result.children.push({
        type: "run" /* Run */,
        ...result.runProps,
        children: [
          { type: "text" /* Text */, text: "\xA0" }
        ]
      });
    }
    return result;
  }
  parseParagraphProperties(elem, paragraph) {
    this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, (c) => {
      if (parseParagraphProperty(c, paragraph, xml_parser_default))
        return true;
      switch (c.localName) {
        case "pStyle":
          paragraph.styleName = xml_parser_default.attr(c, "val");
          break;
        case "cnfStyle":
          paragraph.className = values.classNameOfCnfStyle(c);
          break;
        case "framePr":
          this.parseFrame(c, paragraph);
          break;
        case "rPr":
          break;
        default:
          return false;
      }
      return true;
    });
  }
  parseFrame(node, paragraph) {
    var dropCap = xml_parser_default.attr(node, "dropCap");
    if (dropCap == "drop")
      paragraph.cssStyle["float"] = "left";
  }
  parseHyperlink(node, parent) {
    var result = { type: "hyperlink" /* Hyperlink */, parent, children: [] };
    result.anchor = xml_parser_default.attr(node, "anchor");
    result.id = xml_parser_default.attr(node, "id");
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "r":
          result.children.push(this.parseRun(c, result));
          break;
      }
    }
    return result;
  }
  parseSmartTag(node, parent) {
    var result = { type: "smartTag" /* SmartTag */, parent, children: [] };
    var uri = xml_parser_default.attr(node, "uri");
    var element = xml_parser_default.attr(node, "element");
    if (uri)
      result.uri = uri;
    if (element)
      result.element = element;
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "r":
          result.children.push(this.parseRun(c, result));
          break;
      }
    }
    return result;
  }
  parseRun(node, parent) {
    var result = { type: "run" /* Run */, parent, children: [] };
    for (let c of xml_parser_default.elements(node)) {
      c = this.checkAlternateContent(c);
      switch (c.localName) {
        case "t":
          result.children.push({
            type: "text" /* Text */,
            text: c.textContent
          });
          break;
        case "delText":
          result.children.push({
            type: "deletedText" /* DeletedText */,
            text: c.textContent
          });
          break;
        case "commentReference":
          result.children.push(new WmlCommentReference(xml_parser_default.attr(c, "id")));
          break;
        case "fldSimple":
          result.children.push({
            type: "simpleField" /* SimpleField */,
            instruction: xml_parser_default.attr(c, "instr"),
            lock: xml_parser_default.boolAttr(c, "lock", false),
            dirty: xml_parser_default.boolAttr(c, "dirty", false)
          });
          break;
        case "instrText":
          result.fieldRun = true;
          result.children.push({
            type: "instruction" /* Instruction */,
            text: c.textContent
          });
          break;
        case "fldChar":
          result.fieldRun = true;
          result.children.push({
            type: "complexField" /* ComplexField */,
            charType: xml_parser_default.attr(c, "fldCharType"),
            lock: xml_parser_default.boolAttr(c, "lock", false),
            dirty: xml_parser_default.boolAttr(c, "dirty", false)
          });
          break;
        case "noBreakHyphen":
          result.children.push({ type: "noBreakHyphen" /* NoBreakHyphen */ });
          break;
        case "br":
          result.children.push({
            type: "break" /* Break */,
            break: xml_parser_default.attr(c, "type") || "textWrapping"
          });
          break;
        case "lastRenderedPageBreak":
          result.children.push({
            type: "break" /* Break */,
            break: "lastRenderedPageBreak"
          });
          break;
        case "sym":
          result.children.push({
            type: "symbol" /* Symbol */,
            font: encloseFontFamily(xml_parser_default.attr(c, "font")),
            char: xml_parser_default.attr(c, "char")
          });
          break;
        case "tab":
          result.children.push({ type: "tab" /* Tab */ });
          break;
        case "footnoteReference":
          result.children.push({
            type: "footnoteReference" /* FootnoteReference */,
            id: xml_parser_default.attr(c, "id")
          });
          break;
        case "endnoteReference":
          result.children.push({
            type: "endnoteReference" /* EndnoteReference */,
            id: xml_parser_default.attr(c, "id")
          });
          break;
        case "drawing":
          let d = this.parseDrawing(c);
          if (d)
            result.children = [d];
          break;
        case "pict":
          result.children.push(this.parseVmlPicture(c));
          break;
        case "rPr":
          this.parseRunProperties(c, result);
          break;
      }
    }
    return result;
  }
  parseMathElement(elem) {
    const propsTag = `${elem.localName}Pr`;
    const result = { type: mmlTagMap[elem.localName], children: [] };
    for (const el of xml_parser_default.elements(elem)) {
      const childType = mmlTagMap[el.localName];
      if (childType) {
        result.children.push(this.parseMathElement(el));
      } else if (el.localName == "r") {
        var run = this.parseRun(el);
        run.type = "mmlRun" /* MmlRun */;
        result.children.push(run);
      } else if (el.localName == propsTag) {
        result.props = this.parseMathProperies(el);
      }
    }
    return result;
  }
  parseMathProperies(elem) {
    const result = {};
    for (const el of xml_parser_default.elements(elem)) {
      switch (el.localName) {
        case "chr":
          result.char = xml_parser_default.attr(el, "val");
          break;
        case "vertJc":
          result.verticalJustification = xml_parser_default.attr(el, "val");
          break;
        case "pos":
          result.position = xml_parser_default.attr(el, "val");
          break;
        case "degHide":
          result.hideDegree = xml_parser_default.boolAttr(el, "val");
          break;
        case "begChr":
          result.beginChar = xml_parser_default.attr(el, "val");
          break;
        case "endChr":
          result.endChar = xml_parser_default.attr(el, "val");
          break;
      }
    }
    return result;
  }
  parseRunProperties(elem, run) {
    this.parseDefaultProperties(elem, run.cssStyle = {}, null, (c) => {
      switch (c.localName) {
        case "rStyle":
          run.styleName = xml_parser_default.attr(c, "val");
          break;
        case "vertAlign":
          run.verticalAlign = values.valueOfVertAlign(c, true);
          break;
        default:
          return false;
      }
      return true;
    });
  }
  parseVmlPicture(elem) {
    const result = { type: "vmlPicture" /* VmlPicture */, children: [] };
    for (const el of xml_parser_default.elements(elem)) {
      const child = parseVmlElement(el, this);
      child && result.children.push(child);
    }
    return result;
  }
  checkAlternateContent(elem) {
    if (elem.localName != "AlternateContent")
      return elem;
    var choice = xml_parser_default.element(elem, "Choice");
    if (choice) {
      var requires = xml_parser_default.attr(choice, "Requires");
      var namespaceURI = elem.lookupNamespaceURI(requires);
      if (supportedNamespaceURIs.includes(namespaceURI))
        return choice.firstElementChild;
    }
    return xml_parser_default.element(elem, "Fallback")?.firstElementChild;
  }
  parseDrawing(node) {
    for (var n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "inline":
        case "anchor":
          return this.parseDrawingWrapper(n);
      }
    }
  }
  parseDrawingWrapper(node) {
    var result = { type: "drawing" /* Drawing */, children: [], cssStyle: {} };
    var isAnchor = node.localName == "anchor";
    let wrapType = null;
    let simplePos = xml_parser_default.boolAttr(node, "simplePos");
    let behindDoc = xml_parser_default.boolAttr(node, "behindDoc");
    let posX = { relative: "page", align: "left", offset: "0" };
    let posY = { relative: "page", align: "top", offset: "0" };
    for (var n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "simplePos":
          if (simplePos) {
            posX.offset = xml_parser_default.lengthAttr(n, "x", LengthUsage.Emu);
            posY.offset = xml_parser_default.lengthAttr(n, "y", LengthUsage.Emu);
          }
          break;
        case "extent":
          result.cssStyle["width"] = xml_parser_default.lengthAttr(n, "cx", LengthUsage.Emu);
          result.cssStyle["height"] = xml_parser_default.lengthAttr(n, "cy", LengthUsage.Emu);
          break;
        case "positionH":
        case "positionV":
          if (!simplePos) {
            let pos = n.localName == "positionH" ? posX : posY;
            var alignNode = xml_parser_default.element(n, "align");
            var offsetNode = xml_parser_default.element(n, "posOffset");
            pos.relative = xml_parser_default.attr(n, "relativeFrom") ?? pos.relative;
            if (alignNode)
              pos.align = alignNode.textContent;
            if (offsetNode)
              pos.offset = convertLength(offsetNode.textContent, LengthUsage.Emu);
          }
          break;
        case "wrapTopAndBottom":
          wrapType = "wrapTopAndBottom";
          break;
        case "wrapNone":
          wrapType = "wrapNone";
          break;
        case "graphic":
          var g = this.parseGraphic(n);
          if (g)
            result.children.push(g);
          break;
      }
    }
    if (wrapType == "wrapTopAndBottom") {
      result.cssStyle["display"] = "block";
      if (posX.align) {
        result.cssStyle["text-align"] = posX.align;
        result.cssStyle["width"] = "100%";
      }
    } else if (wrapType == "wrapNone") {
      result.cssStyle["display"] = "block";
      result.cssStyle["position"] = "relative";
      result.cssStyle["width"] = "0px";
      result.cssStyle["height"] = "0px";
      if (posX.offset)
        result.cssStyle["left"] = posX.offset;
      if (posY.offset)
        result.cssStyle["top"] = posY.offset;
    } else if (isAnchor && (posX.align == "left" || posX.align == "right")) {
      result.cssStyle["float"] = posX.align;
    }
    return result;
  }
  parseGraphic(elem) {
    var graphicData = xml_parser_default.element(elem, "graphicData");
    for (let n of xml_parser_default.elements(graphicData)) {
      switch (n.localName) {
        case "pic":
          return this.parsePicture(n);
      }
    }
    return null;
  }
  parsePicture(elem) {
    var result = { type: "image" /* Image */, src: "", cssStyle: {} };
    var blipFill = xml_parser_default.element(elem, "blipFill");
    var blip = xml_parser_default.element(blipFill, "blip");
    var srcRect = xml_parser_default.element(blipFill, "srcRect");
    result.src = xml_parser_default.attr(blip, "embed");
    if (srcRect) {
      result.srcRect = [
        xml_parser_default.intAttr(srcRect, "l", 0) / 1e5,
        xml_parser_default.intAttr(srcRect, "t", 0) / 1e5,
        xml_parser_default.intAttr(srcRect, "r", 0) / 1e5,
        xml_parser_default.intAttr(srcRect, "b", 0) / 1e5
      ];
    }
    var spPr = xml_parser_default.element(elem, "spPr");
    var xfrm = xml_parser_default.element(spPr, "xfrm");
    result.cssStyle["position"] = "relative";
    if (xfrm) {
      result.rotation = xml_parser_default.intAttr(xfrm, "rot", 0) / 6e4;
      for (var n of xml_parser_default.elements(xfrm)) {
        switch (n.localName) {
          case "ext":
            result.cssStyle["width"] = xml_parser_default.lengthAttr(n, "cx", LengthUsage.Emu);
            result.cssStyle["height"] = xml_parser_default.lengthAttr(n, "cy", LengthUsage.Emu);
            break;
          case "off":
            result.cssStyle["left"] = xml_parser_default.lengthAttr(n, "x", LengthUsage.Emu);
            result.cssStyle["top"] = xml_parser_default.lengthAttr(n, "y", LengthUsage.Emu);
            break;
        }
      }
    }
    return result;
  }
  parseTable(node) {
    var result = { type: "table" /* Table */, children: [] };
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "tr":
          result.children.push(this.parseTableRow(c));
          break;
        case "tblGrid":
          result.columns = this.parseTableColumns(c);
          break;
        case "tblPr":
          this.parseTableProperties(c, result);
          break;
      }
    }
    return result;
  }
  parseTableColumns(node) {
    var result = [];
    for (const n of xml_parser_default.elements(node)) {
      switch (n.localName) {
        case "gridCol":
          result.push({ width: xml_parser_default.lengthAttr(n, "w") });
          break;
      }
    }
    return result;
  }
  parseTableProperties(elem, table) {
    table.cssStyle = {};
    table.cellStyle = {};
    this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, (c) => {
      switch (c.localName) {
        case "tblStyle":
          table.styleName = xml_parser_default.attr(c, "val");
          break;
        case "tblLook":
          table.className = values.classNameOftblLook(c);
          break;
        case "tblpPr":
          this.parseTablePosition(c, table);
          break;
        case "tblStyleColBandSize":
          table.colBandSize = xml_parser_default.intAttr(c, "val");
          break;
        case "tblStyleRowBandSize":
          table.rowBandSize = xml_parser_default.intAttr(c, "val");
          break;
        case "hidden":
          table.cssStyle["display"] = "none";
          break;
        default:
          return false;
      }
      return true;
    });
    switch (table.cssStyle["text-align"]) {
      case "center":
        delete table.cssStyle["text-align"];
        table.cssStyle["margin-left"] = "auto";
        table.cssStyle["margin-right"] = "auto";
        break;
      case "right":
        delete table.cssStyle["text-align"];
        table.cssStyle["margin-left"] = "auto";
        break;
    }
  }
  parseTablePosition(node, table) {
    var topFromText = xml_parser_default.lengthAttr(node, "topFromText");
    var bottomFromText = xml_parser_default.lengthAttr(node, "bottomFromText");
    var rightFromText = xml_parser_default.lengthAttr(node, "rightFromText");
    var leftFromText = xml_parser_default.lengthAttr(node, "leftFromText");
    table.cssStyle["float"] = "left";
    table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
    table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
    table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
    table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
  }
  parseTableRow(node) {
    var result = { type: "row" /* Row */, children: [] };
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "tc":
          result.children.push(this.parseTableCell(c));
          break;
        case "trPr":
        case "tblPrEx":
          this.parseTableRowProperties(c, result);
          break;
      }
    }
    return result;
  }
  parseTableRowProperties(elem, row) {
    row.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case "cnfStyle":
          row.className = values.classNameOfCnfStyle(c);
          break;
        case "tblHeader":
          row.isHeader = xml_parser_default.boolAttr(c, "val");
          break;
        case "gridBefore":
          row.gridBefore = xml_parser_default.intAttr(c, "val");
          break;
        case "gridAfter":
          row.gridAfter = xml_parser_default.intAttr(c, "val");
          break;
        default:
          return false;
      }
      return true;
    });
  }
  parseTableCell(node) {
    var result = { type: "cell" /* Cell */, children: [] };
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "tbl":
          result.children.push(this.parseTable(c));
          break;
        case "p":
          result.children.push(this.parseParagraph(c));
          break;
        case "tcPr":
          this.parseTableCellProperties(c, result);
          break;
      }
    }
    return result;
  }
  parseTableCellProperties(elem, cell) {
    cell.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case "gridSpan":
          cell.span = xml_parser_default.intAttr(c, "val", null);
          break;
        case "vMerge":
          cell.verticalMerge = xml_parser_default.attr(c, "val") ?? "continue";
          break;
        case "cnfStyle":
          cell.className = values.classNameOfCnfStyle(c);
          break;
        default:
          return false;
      }
      return true;
    });
    this.parseTableCellVerticalText(elem, cell);
  }
  parseTableCellVerticalText(elem, cell) {
    const directionMap = {
      "btLr": {
        writingMode: "vertical-rl",
        transform: "rotate(180deg)"
      },
      "lrTb": {
        writingMode: "vertical-lr",
        transform: "none"
      },
      "tbRl": {
        writingMode: "vertical-rl",
        transform: "none"
      }
    };
    for (const c of xml_parser_default.elements(elem)) {
      if (c.localName === "textDirection") {
        const direction = xml_parser_default.attr(c, "val");
        const style = directionMap[direction] || { writingMode: "horizontal-tb" };
        cell.cssStyle["writing-mode"] = style.writingMode;
        cell.cssStyle["transform"] = style.transform;
      }
    }
  }
  parseDefaultProperties(elem, style = null, childStyle = null, handler = null) {
    style = style || {};
    for (const c of xml_parser_default.elements(elem)) {
      if (handler?.(c))
        continue;
      switch (c.localName) {
        case "jc":
          style["text-align"] = values.valueOfJc(c);
          break;
        case "textAlignment":
          style["vertical-align"] = values.valueOfTextAlignment(c);
          break;
        case "color":
          style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
          break;
        case "sz":
          style["font-size"] = style["min-height"] = xml_parser_default.lengthAttr(c, "val", LengthUsage.FontSize);
          break;
        case "shd":
          style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
          break;
        case "highlight":
          style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
          break;
        case "vertAlign":
          break;
        case "position":
          style.verticalAlign = xml_parser_default.lengthAttr(c, "val", LengthUsage.FontSize);
          break;
        case "tcW":
          if (this.options.ignoreWidth)
            break;
        case "tblW":
          style["width"] = values.valueOfSize(c, "w");
          break;
        case "trHeight":
          this.parseTrHeight(c, style);
          break;
        case "strike":
          style["text-decoration"] = xml_parser_default.boolAttr(c, "val", true) ? "line-through" : "none";
          break;
        case "b":
          style["font-weight"] = xml_parser_default.boolAttr(c, "val", true) ? "bold" : "normal";
          break;
        case "i":
          style["font-style"] = xml_parser_default.boolAttr(c, "val", true) ? "italic" : "normal";
          break;
        case "caps":
          style["text-transform"] = xml_parser_default.boolAttr(c, "val", true) ? "uppercase" : "none";
          break;
        case "smallCaps":
          style["font-variant"] = xml_parser_default.boolAttr(c, "val", true) ? "small-caps" : "none";
          break;
        case "u":
          this.parseUnderline(c, style);
          break;
        case "ind":
        case "tblInd":
          this.parseIndentation(c, style);
          break;
        case "rFonts":
          this.parseFont(c, style);
          break;
        case "tblBorders":
          this.parseBorderProperties(c, childStyle || style);
          break;
        case "tblCellSpacing":
          style["border-spacing"] = values.valueOfMargin(c);
          style["border-collapse"] = "separate";
          break;
        case "pBdr":
          this.parseBorderProperties(c, style);
          break;
        case "bdr":
          style["border"] = values.valueOfBorder(c);
          break;
        case "tcBorders":
          this.parseBorderProperties(c, style);
          break;
        case "vanish":
          if (xml_parser_default.boolAttr(c, "val", true))
            style["display"] = "none";
          break;
        case "kern":
          break;
        case "noWrap":
          break;
        case "tblCellMar":
        case "tcMar":
          this.parseMarginProperties(c, childStyle || style);
          break;
        case "tblLayout":
          style["table-layout"] = values.valueOfTblLayout(c);
          break;
        case "vAlign":
          style["vertical-align"] = values.valueOfTextAlignment(c);
          break;
        case "spacing":
          if (elem.localName == "pPr")
            this.parseSpacing(c, style);
          break;
        case "wordWrap":
          if (xml_parser_default.boolAttr(c, "val"))
            style["overflow-wrap"] = "break-word";
          break;
        case "suppressAutoHyphens":
          style["hyphens"] = xml_parser_default.boolAttr(c, "val", true) ? "none" : "auto";
          break;
        case "lang":
          style["$lang"] = xml_parser_default.attr(c, "val");
          break;
        case "rtl":
        case "bidi":
          if (xml_parser_default.boolAttr(c, "val", true))
            style["direction"] = "rtl";
          break;
        case "bCs":
        case "iCs":
        case "szCs":
        case "tabs":
        //ignore - tabs is parsed by other parser
        case "outlineLvl":
        //TODO
        case "contextualSpacing":
        //TODO
        case "tblStyleColBandSize":
        //TODO
        case "tblStyleRowBandSize":
        //TODO
        case "webHidden":
        //TODO - maybe web-hidden should be implemented
        case "pageBreakBefore":
        //TODO - maybe ignore 
        case "suppressLineNumbers":
        //TODO - maybe ignore
        case "keepLines":
        //TODO - maybe ignore
        case "keepNext":
        //TODO - maybe ignore
        case "widowControl":
        //TODO - maybe ignore 
        case "noProof":
          break;
        default:
          if (this.options.debug)
            console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
          break;
      }
    }
    return style;
  }
  parseUnderline(node, style) {
    var val = xml_parser_default.attr(node, "val");
    if (val == null)
      return;
    switch (val) {
      case "dash":
      case "dashDotDotHeavy":
      case "dashDotHeavy":
      case "dashedHeavy":
      case "dashLong":
      case "dashLongHeavy":
      case "dotDash":
      case "dotDotDash":
        style["text-decoration"] = "underline dashed";
        break;
      case "dotted":
      case "dottedHeavy":
        style["text-decoration"] = "underline dotted";
        break;
      case "double":
        style["text-decoration"] = "underline double";
        break;
      case "single":
      case "thick":
        style["text-decoration"] = "underline";
        break;
      case "wave":
      case "wavyDouble":
      case "wavyHeavy":
        style["text-decoration"] = "underline wavy";
        break;
      case "words":
        style["text-decoration"] = "underline";
        break;
      case "none":
        style["text-decoration"] = "none";
        break;
    }
    var col = xmlUtil.colorAttr(node, "color");
    if (col)
      style["text-decoration-color"] = col;
  }
  parseFont(node, style) {
    var ascii = xml_parser_default.attr(node, "ascii");
    var asciiTheme = values.themeValue(node, "asciiTheme");
    var eastAsia = xml_parser_default.attr(node, "eastAsia");
    var fonts = [ascii, asciiTheme, eastAsia].filter((x) => x).map((x) => encloseFontFamily(x));
    if (fonts.length > 0)
      style["font-family"] = [...new Set(fonts)].join(", ");
  }
  parseIndentation(node, style) {
    var firstLine = xml_parser_default.lengthAttr(node, "firstLine");
    var hanging = xml_parser_default.lengthAttr(node, "hanging");
    var left = xml_parser_default.lengthAttr(node, "left");
    var start = xml_parser_default.lengthAttr(node, "start");
    var right = xml_parser_default.lengthAttr(node, "right");
    var end = xml_parser_default.lengthAttr(node, "end");
    if (firstLine) style["text-indent"] = firstLine;
    if (hanging) style["text-indent"] = `-${hanging}`;
    if (left || start) style["margin-inline-start"] = left || start;
    if (right || end) style["margin-inline-end"] = right || end;
  }
  parseSpacing(node, style) {
    var before = xml_parser_default.lengthAttr(node, "before");
    var after = xml_parser_default.lengthAttr(node, "after");
    var line = xml_parser_default.intAttr(node, "line", null);
    var lineRule = xml_parser_default.attr(node, "lineRule");
    if (before) style["margin-top"] = before;
    if (after) style["margin-bottom"] = after;
    if (line !== null) {
      switch (lineRule) {
        case "auto":
          style["line-height"] = `${(line / 240).toFixed(2)}`;
          break;
        case "atLeast":
          style["line-height"] = `calc(100% + ${line / 20}pt)`;
          break;
        default:
          style["line-height"] = style["min-height"] = `${line / 20}pt`;
          break;
      }
    }
  }
  parseMarginProperties(node, output) {
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "left":
          output["padding-left"] = values.valueOfMargin(c);
          break;
        case "right":
          output["padding-right"] = values.valueOfMargin(c);
          break;
        case "top":
          output["padding-top"] = values.valueOfMargin(c);
          break;
        case "bottom":
          output["padding-bottom"] = values.valueOfMargin(c);
          break;
      }
    }
  }
  parseTrHeight(node, output) {
    switch (xml_parser_default.attr(node, "hRule")) {
      case "exact":
        output["height"] = xml_parser_default.lengthAttr(node, "val");
        break;
      case "atLeast":
      default:
        output["height"] = xml_parser_default.lengthAttr(node, "val");
        break;
    }
  }
  parseBorderProperties(node, output) {
    for (const c of xml_parser_default.elements(node)) {
      switch (c.localName) {
        case "start":
        case "left":
          output["border-left"] = values.valueOfBorder(c);
          break;
        case "end":
        case "right":
          output["border-right"] = values.valueOfBorder(c);
          break;
        case "top":
          output["border-top"] = values.valueOfBorder(c);
          break;
        case "bottom":
          output["border-bottom"] = values.valueOfBorder(c);
          break;
      }
    }
  }
};
var knownColors = ["black", "blue", "cyan", "darkBlue", "darkCyan", "darkGray", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "green", "lightGray", "magenta", "none", "red", "white", "yellow"];
var xmlUtil = class {
  static colorAttr(node, attrName, defValue = null, autoColor = "black") {
    var v = xml_parser_default.attr(node, attrName);
    if (v) {
      if (v == "auto") {
        return autoColor;
      } else if (knownColors.includes(v)) {
        return v;
      }
      return `#${v}`;
    }
    var themeColor = xml_parser_default.attr(node, "themeColor");
    return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
  }
};
var values = class _values {
  static themeValue(c, attr) {
    var val = xml_parser_default.attr(c, attr);
    return val ? `var(--docx-${val}-font)` : null;
  }
  static valueOfSize(c, attr) {
    var type = LengthUsage.Dxa;
    switch (xml_parser_default.attr(c, "type")) {
      case "dxa":
        break;
      case "pct":
        type = LengthUsage.Percent;
        break;
      case "auto":
        return "auto";
    }
    return xml_parser_default.lengthAttr(c, attr, type);
  }
  static valueOfMargin(c) {
    return xml_parser_default.lengthAttr(c, "w");
  }
  static valueOfBorder(c) {
    var type = _values.parseBorderType(xml_parser_default.attr(c, "val"));
    if (type == "none")
      return "none";
    var color = xmlUtil.colorAttr(c, "color");
    var size = xml_parser_default.lengthAttr(c, "sz", LengthUsage.Border);
    return `${size} ${type} ${color == "auto" ? autos.borderColor : color}`;
  }
  static parseBorderType(type) {
    switch (type) {
      case "single":
        return "solid";
      case "dashDotStroked":
        return "solid";
      case "dashed":
        return "dashed";
      case "dashSmallGap":
        return "dashed";
      case "dotDash":
        return "dotted";
      case "dotDotDash":
        return "dotted";
      case "dotted":
        return "dotted";
      case "double":
        return "double";
      case "doubleWave":
        return "double";
      case "inset":
        return "inset";
      case "nil":
        return "none";
      case "none":
        return "none";
      case "outset":
        return "outset";
      case "thick":
        return "solid";
      case "thickThinLargeGap":
        return "solid";
      case "thickThinMediumGap":
        return "solid";
      case "thickThinSmallGap":
        return "solid";
      case "thinThickLargeGap":
        return "solid";
      case "thinThickMediumGap":
        return "solid";
      case "thinThickSmallGap":
        return "solid";
      case "thinThickThinLargeGap":
        return "solid";
      case "thinThickThinMediumGap":
        return "solid";
      case "thinThickThinSmallGap":
        return "solid";
      case "threeDEmboss":
        return "solid";
      case "threeDEngrave":
        return "solid";
      case "triple":
        return "double";
      case "wave":
        return "solid";
    }
    return "solid";
  }
  static valueOfTblLayout(c) {
    var type = xml_parser_default.attr(c, "val");
    return type == "fixed" ? "fixed" : "auto";
  }
  static classNameOfCnfStyle(c) {
    const val = xml_parser_default.attr(c, "val");
    const classes = [
      "first-row",
      "last-row",
      "first-col",
      "last-col",
      "odd-col",
      "even-col",
      "odd-row",
      "even-row",
      "ne-cell",
      "nw-cell",
      "se-cell",
      "sw-cell"
    ];
    return classes.filter((_, i) => val[i] == "1").join(" ");
  }
  static valueOfJc(c) {
    var type = xml_parser_default.attr(c, "val");
    switch (type) {
      case "start":
      case "left":
        return "left";
      case "center":
        return "center";
      case "end":
      case "right":
        return "right";
      case "both":
        return "justify";
    }
    return type;
  }
  static valueOfVertAlign(c, asTagName = false) {
    var type = xml_parser_default.attr(c, "val");
    switch (type) {
      case "subscript":
        return "sub";
      case "superscript":
        return asTagName ? "sup" : "super";
    }
    return asTagName ? null : type;
  }
  static valueOfTextAlignment(c) {
    var type = xml_parser_default.attr(c, "val");
    switch (type) {
      case "auto":
      case "baseline":
        return "baseline";
      case "top":
        return "top";
      case "center":
        return "middle";
      case "bottom":
        return "bottom";
    }
    return type;
  }
  static addSize(a, b) {
    if (a == null) return b;
    if (b == null) return a;
    return `calc(${a} + ${b})`;
  }
  static classNameOftblLook(c) {
    const val = xml_parser_default.hexAttr(c, "val", 0);
    let className = "";
    if (xml_parser_default.boolAttr(c, "firstRow") || val & 32) className += " first-row";
    if (xml_parser_default.boolAttr(c, "lastRow") || val & 64) className += " last-row";
    if (xml_parser_default.boolAttr(c, "firstColumn") || val & 128) className += " first-col";
    if (xml_parser_default.boolAttr(c, "lastColumn") || val & 256) className += " last-col";
    if (xml_parser_default.boolAttr(c, "noHBand") || val & 512) className += " no-hband";
    if (xml_parser_default.boolAttr(c, "noVBand") || val & 1024) className += " no-vband";
    return className.trim();
  }
};

// src/javascript.ts
var defaultTab = { pos: 0, leader: "none", style: "left" };
var maxTabs = 50;
function computePixelToPoint(container = document.body) {
  const temp = document.createElement("div");
  temp.style.width = "100pt";
  container.appendChild(temp);
  const result = 100 / temp.offsetWidth;
  container.removeChild(temp);
  return result;
}
function updateTabStop(elem, tabs, defaultTabSize, pixelToPoint = 72 / 96) {
  const p = elem.closest("p");
  const ebb = elem.getBoundingClientRect();
  const pbb = p.getBoundingClientRect();
  const pcs = getComputedStyle(p);
  const tabStops = tabs?.length > 0 ? tabs.map((t) => ({
    pos: lengthToPoint(t.position),
    leader: t.leader,
    style: t.style
  })).sort((a, b) => a.pos - b.pos) : [defaultTab];
  const lastTab = tabStops[tabStops.length - 1];
  const pWidthPt = pbb.width * pixelToPoint;
  const size = lengthToPoint(defaultTabSize);
  let pos = lastTab.pos + size;
  if (pos < pWidthPt) {
    for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
      tabStops.push({ ...defaultTab, pos });
    }
  }
  const marginLeft = parseFloat(pcs.marginLeft);
  const pOffset = pbb.left + marginLeft;
  const left = (ebb.left - pOffset) * pixelToPoint;
  const tab = tabStops.find((t) => t.style != "clear" && t.pos > left);
  if (tab == null)
    return;
  let width = 1;
  if (tab.style == "right" || tab.style == "center") {
    const tabStops2 = Array.from(p.querySelectorAll(`.${elem.className}`));
    const nextIdx = tabStops2.indexOf(elem) + 1;
    const range = document.createRange();
    range.setStart(elem, 1);
    if (nextIdx < tabStops2.length) {
      range.setEndBefore(tabStops2[nextIdx]);
    } else {
      range.setEndAfter(p);
    }
    const mul = tab.style == "center" ? 0.5 : 1;
    const nextBB = range.getBoundingClientRect();
    const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);
    width = tab.pos - offset * pixelToPoint;
  } else {
    width = tab.pos - left;
  }
  elem.innerHTML = "&nbsp;";
  elem.style.textDecoration = "inherit";
  elem.style.wordSpacing = `${width.toFixed(0)}pt`;
  switch (tab.leader) {
    case "dot":
    case "middleDot":
      elem.style.textDecoration = "underline";
      elem.style.textDecorationStyle = "dotted";
      break;
    case "hyphen":
    case "heavy":
    case "underscore":
      elem.style.textDecoration = "underline";
      break;
  }
}
function lengthToPoint(length) {
  return parseFloat(length);
}

// src/html-renderer.ts
var ns2 = {
  svg: "http://www.w3.org/2000/svg",
  mathML: "http://www.w3.org/1998/Math/MathML"
};
var HtmlRenderer = class {
  constructor(htmlDocument) {
    this.htmlDocument = htmlDocument;
    this.className = "docx";
    this.styleMap = {};
    this.currentPart = null;
    this.tableVerticalMerges = [];
    this.currentVerticalMerge = null;
    this.tableCellPositions = [];
    this.currentCellPosition = null;
    this.footnoteMap = {};
    this.endnoteMap = {};
    this.currentEndnoteIds = [];
    this.usedHederFooterParts = [];
    this.currentTabs = [];
    this.commentMap = {};
    this.tasks = [];
    this.postRenderTasks = [];
    this.numFormatMapping = {
      none: "none",
      bullet: "disc",
      decimal: "decimal",
      lowerLetter: "lower-alpha",
      upperLetter: "upper-alpha",
      lowerRoman: "lower-roman",
      upperRoman: "upper-roman",
      decimalZero: "decimal-leading-zero",
      // 01,02,03,...
      // ordinal: "", // 1st, 2nd, 3rd,...
      // ordinalText: "", //First, Second, Third, ...
      // cardinalText: "", //One,Two Three,...
      // numberInDash: "", //-1-,-2-,-3-, ...
      // hex: "upper-hexadecimal",
      aiueo: "katakana",
      aiueoFullWidth: "katakana",
      chineseCounting: "simp-chinese-informal",
      chineseCountingThousand: "simp-chinese-informal",
      chineseLegalSimplified: "simp-chinese-formal",
      // 中文大写
      chosung: "hangul-consonant",
      ideographDigital: "cjk-ideographic",
      ideographTraditional: "cjk-heavenly-stem",
      // 十天干
      ideographLegalTraditional: "trad-chinese-formal",
      ideographZodiac: "cjk-earthly-branch",
      // 十二地支
      iroha: "katakana-iroha",
      irohaFullWidth: "katakana-iroha",
      japaneseCounting: "japanese-informal",
      japaneseDigitalTenThousand: "cjk-decimal",
      japaneseLegal: "japanese-formal",
      thaiNumbers: "thai",
      koreanCounting: "korean-hangul-formal",
      koreanDigital: "korean-hangul-formal",
      koreanDigital2: "korean-hanja-informal",
      hebrew1: "hebrew",
      hebrew2: "hebrew",
      hindiNumbers: "devanagari",
      ganada: "hangul",
      taiwaneseCounting: "cjk-ideographic",
      taiwaneseCountingThousand: "cjk-ideographic",
      taiwaneseDigital: "cjk-decimal"
    };
  }
  async render(document2, bodyContainer, styleContainer = null, options) {
    this.document = document2;
    this.options = options;
    this.className = options.className;
    this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ":root";
    this.styleMap = null;
    this.tasks = [];
    if (this.options.renderComments && globalThis.Highlight) {
      this.commentHighlight = new Highlight();
    }
    styleContainer = styleContainer || bodyContainer;
    removeAllElements(styleContainer);
    removeAllElements(bodyContainer);
    styleContainer.appendChild(this.createComment("docxjs library predefined styles"));
    styleContainer.appendChild(this.renderDefaultStyle());
    if (document2.themePart) {
      styleContainer.appendChild(this.createComment("docxjs document theme values"));
      this.renderTheme(document2.themePart, styleContainer);
    }
    if (document2.stylesPart != null) {
      this.styleMap = this.processStyles(document2.stylesPart.styles);
      styleContainer.appendChild(this.createComment("docxjs document styles"));
      styleContainer.appendChild(this.renderStyles(document2.stylesPart.styles));
    }
    if (document2.numberingPart) {
      this.processNumberings(document2.numberingPart.domNumberings);
      styleContainer.appendChild(this.createComment("docxjs document numbering styles"));
      styleContainer.appendChild(this.renderNumbering(document2.numberingPart.domNumberings, styleContainer));
    }
    if (document2.footnotesPart) {
      this.footnoteMap = keyBy(document2.footnotesPart.notes, (x) => x.id);
    }
    if (document2.endnotesPart) {
      this.endnoteMap = keyBy(document2.endnotesPart.notes, (x) => x.id);
    }
    if (document2.settingsPart) {
      this.defaultTabSize = document2.settingsPart.settings?.defaultTabStop;
    }
    if (!options.ignoreFonts && document2.fontTablePart)
      this.renderFontTable(document2.fontTablePart, styleContainer);
    var sectionElements = this.renderSections(document2.documentPart.body);
    if (this.options.inWrapper) {
      bodyContainer.appendChild(this.renderWrapper(sectionElements));
    } else {
      appendChildren(bodyContainer, sectionElements);
    }
    if (this.commentHighlight && options.renderComments) {
      CSS.highlights.set(`${this.className}-comments`, this.commentHighlight);
    }
    this.postRenderTasks.forEach((t) => t());
    await Promise.allSettled(this.tasks);
    this.refreshTabStops();
  }
  renderTheme(themePart, styleContainer) {
    const variables = {};
    const fontScheme = themePart.theme?.fontScheme;
    if (fontScheme) {
      if (fontScheme.majorFont) {
        variables["--docx-majorHAnsi-font"] = fontScheme.majorFont.latinTypeface;
      }
      if (fontScheme.minorFont) {
        variables["--docx-minorHAnsi-font"] = fontScheme.minorFont.latinTypeface;
      }
    }
    const colorScheme = themePart.theme?.colorScheme;
    if (colorScheme) {
      for (let [k, v] of Object.entries(colorScheme.colors)) {
        variables[`--docx-${k}-color`] = `#${v}`;
      }
    }
    const cssText = this.styleToString(`.${this.className}`, variables);
    styleContainer.appendChild(this.createStyleElement(cssText));
  }
  renderFontTable(fontsPart, styleContainer) {
    const fonts = [];
    for (let f of fontsPart.fonts) {
      for (let ref of f.embedFontRefs) {
        this.tasks.push(this.document.loadFont(ref.id, ref.key).then((fontData) => {
          const cssValues = {
            "font-family": encloseFontFamily(f.name),
            "src": `url(${fontData})`
          };
          if (ref.type == "bold" || ref.type == "boldItalic") {
            cssValues["font-weight"] = "bold";
          }
          if (ref.type == "italic" || ref.type == "boldItalic") {
            cssValues["font-style"] = "italic";
          }
          const cssText = this.styleToString("@font-face", cssValues);
          styleContainer.appendChild(this.createComment(`docxjs ${f.name} font`));
          styleContainer.appendChild(this.createStyleElement(cssText));
        }));
      }
    }
  }
  processStyleName(className) {
    return className ? `${this.className}_${escapeClassName(className)}` : this.className;
  }
  processStyles(styles) {
    const stylesMap = keyBy(styles.filter((x) => x.id != null), (x) => x.id);
    for (const style of styles.filter((x) => x.basedOn)) {
      var baseStyle = stylesMap[style.basedOn];
      if (baseStyle) {
        style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
        style.runProps = mergeDeep(style.runProps, baseStyle.runProps);
        for (const baseValues of baseStyle.styles) {
          const styleValues = style.styles.find((x) => x.target == baseValues.target);
          if (styleValues) {
            this.copyStyleProperties(baseValues.values, styleValues.values);
          } else {
            style.styles.push({ ...baseValues, values: { ...baseValues.values } });
          }
        }
      } else if (this.options.debug)
        console.warn(`Can't find base style ${style.basedOn}`);
    }
    for (let style of styles) {
      style.cssName = this.processStyleName(style.id);
    }
    return stylesMap;
  }
  processNumberings(numberings) {
    for (let num of numberings.filter((n) => n.pStyleName)) {
      const style = this.findStyle(num.pStyleName);
      if (style?.paragraphProps?.numbering) {
        style.paragraphProps.numbering.level = num.level;
      }
    }
  }
  processElement(element) {
    if (element.children) {
      for (var e of element.children) {
        e.parent = element;
        if (e.type == "table" /* Table */) {
          this.processTable(e);
        } else {
          this.processElement(e);
        }
      }
    }
  }
  processTable(table) {
    for (var r of table.children) {
      for (var c of r.children) {
        c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
          "border-left",
          "border-right",
          "border-top",
          "border-bottom",
          "padding-left",
          "padding-right",
          "padding-top",
          "padding-bottom"
        ]);
        this.processElement(c);
      }
    }
  }
  copyStyleProperties(input, output, attrs = null) {
    if (!input)
      return output;
    if (output == null) output = {};
    if (attrs == null) attrs = Object.getOwnPropertyNames(input);
    for (var key of attrs) {
      if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
        output[key] = input[key];
    }
    return output;
  }
  createPageElement(className, props) {
    var elem = this.createElement("section", { className });
    if (props) {
      if (props.pageMargins) {
        elem.style.paddingLeft = props.pageMargins.left;
        elem.style.paddingRight = props.pageMargins.right;
        elem.style.paddingTop = props.pageMargins.top;
        elem.style.paddingBottom = props.pageMargins.bottom;
      }
      if (props.pageSize) {
        if (!this.options.ignoreWidth)
          elem.style.width = props.pageSize.width;
        if (!this.options.ignoreHeight)
          elem.style.minHeight = props.pageSize.height;
      }
    }
    return elem;
  }
  createSectionContent(props) {
    var elem = this.createElement("article");
    if (props.columns && props.columns.numberOfColumns) {
      elem.style.columnCount = `${props.columns.numberOfColumns}`;
      elem.style.columnGap = props.columns.space;
      if (props.columns.separator) {
        elem.style.columnRule = "1px solid black";
      }
    }
    return elem;
  }
  renderSections(document2) {
    const result = [];
    this.processElement(document2);
    const sections = this.splitBySection(document2.children, document2.props);
    const pages = this.groupByPageBreaks(sections);
    let prevProps = null;
    for (let i = 0, l = pages.length; i < l; i++) {
      this.currentFootnoteIds = [];
      const section = pages[i][0];
      let props = section.sectProps;
      const pageElement = this.createPageElement(this.className, props);
      this.renderStyleValues(document2.cssStyle, pageElement);
      this.options.renderHeaders && this.renderHeaderFooter(
        props.headerRefs,
        props,
        result.length,
        prevProps != props,
        pageElement
      );
      for (const sect of pages[i]) {
        var contentElement = this.createSectionContent(sect.sectProps);
        this.renderElements(sect.elements, contentElement);
        pageElement.appendChild(contentElement);
        props = sect.sectProps;
      }
      if (this.options.renderFootnotes) {
        this.renderNotes(this.currentFootnoteIds, this.footnoteMap, pageElement);
      }
      if (this.options.renderEndnotes && i == l - 1) {
        this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
      }
      this.options.renderFooters && this.renderHeaderFooter(
        props.footerRefs,
        props,
        result.length,
        prevProps != props,
        pageElement
      );
      result.push(pageElement);
      prevProps = props;
    }
    return result;
  }
  renderHeaderFooter(refs, props, page, firstOfSection, into) {
    if (!refs) return;
    var ref = (props.titlePage && firstOfSection ? refs.find((x) => x.type == "first") : null) ?? (page % 2 == 1 ? refs.find((x) => x.type == "even") : null) ?? refs.find((x) => x.type == "default");
    var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart);
    if (part) {
      this.currentPart = part;
      if (!this.usedHederFooterParts.includes(part.path)) {
        this.processElement(part.rootElement);
        this.usedHederFooterParts.push(part.path);
      }
      const [el] = this.renderElements([part.rootElement], into);
      if (props?.pageMargins) {
        if (part.rootElement.type === "header" /* Header */) {
          el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
          el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
        } else if (part.rootElement.type === "footer" /* Footer */) {
          el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
          el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
        }
      }
      this.currentPart = null;
    }
  }
  isPageBreakElement(elem) {
    if (elem.type != "break" /* Break */)
      return false;
    if (elem.break == "lastRenderedPageBreak")
      return !this.options.ignoreLastRenderedPageBreak;
    return elem.break == "page";
  }
  isPageBreakSection(prev, next) {
    if (!prev) return false;
    if (!next) return false;
    return prev.pageSize?.orientation != next.pageSize?.orientation || prev.pageSize?.width != next.pageSize?.width || prev.pageSize?.height != next.pageSize?.height;
  }
  splitBySection(elements, defaultProps) {
    var current = { sectProps: null, elements: [], pageBreak: false };
    var result = [current];
    for (let elem of elements) {
      if (elem.type == "paragraph" /* Paragraph */) {
        const s = this.findStyle(elem.styleName);
        if (s?.paragraphProps?.pageBreakBefore) {
          current.sectProps = sectProps;
          current.pageBreak = true;
          current = { sectProps: null, elements: [], pageBreak: false };
          result.push(current);
        }
      }
      current.elements.push(elem);
      if (elem.type == "paragraph" /* Paragraph */) {
        const p = elem;
        var sectProps = p.sectionProps;
        var pBreakIndex = -1;
        var rBreakIndex = -1;
        if (this.options.breakPages && p.children) {
          pBreakIndex = p.children.findIndex((r) => {
            rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
            return rBreakIndex != -1;
          });
        }
        if (sectProps || pBreakIndex != -1) {
          current.sectProps = sectProps;
          current.pageBreak = pBreakIndex != -1;
          current = { sectProps: null, elements: [], pageBreak: false };
          result.push(current);
        }
        if (pBreakIndex != -1) {
          let breakRun = p.children[pBreakIndex];
          let splitRun = rBreakIndex < breakRun.children.length - 1;
          if (pBreakIndex < p.children.length - 1 || splitRun) {
            var children = elem.children;
            var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
            elem.children = children.slice(0, pBreakIndex);
            current.elements.push(newParagraph);
            if (splitRun) {
              let runChildren = breakRun.children;
              let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
              elem.children.push(newRun);
              breakRun.children = runChildren.slice(rBreakIndex);
            }
          }
        }
      }
    }
    let currentSectProps = null;
    for (let i = result.length - 1; i >= 0; i--) {
      if (result[i].sectProps == null) {
        result[i].sectProps = currentSectProps ?? defaultProps;
      } else {
        currentSectProps = result[i].sectProps;
      }
    }
    return result;
  }
  groupByPageBreaks(sections) {
    let current = [];
    let prev;
    const result = [current];
    for (let s of sections) {
      current.push(s);
      if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
        result.push(current = []);
      prev = s.sectProps;
    }
    return result.filter((x) => x.length > 0);
  }
  renderWrapper(children) {
    return this.createElement("div", { className: `${this.className}-wrapper` }, children);
  }
  renderDefaultStyle() {
    var c = this.className;
    var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
    if (this.options.hideWrapperOnPrint) {
      wrapperStyle = `@media not print { ${wrapperStyle} }`;
    }
    var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; vertical-align: text-bottom; }
.${c} *, .${c} *:before { vertical-align: inherit; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
.${c} .textbox p { font-size: 10pt; line-height: 1.1; min-height: initial; }
`;
    if (this.options.renderComments) {
      styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`;
    }
    ;
    return this.createStyleElement(styleText);
  }
  // renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
  //     let css = "";
  //     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
  //     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
  //     const topCounters = [];
  //     for(let num of numberingPart.numberings) {
  //         const absNum = numberingMap[num.abstractId];
  //         for(let lvl of absNum.levels) {
  //             const className = this.numberingClass(num.id, lvl.level);
  //             let listStyleType = "none";
  //             if(lvl.text && lvl.format == 'decimal') {
  //                 const counter = this.numberingCounter(num.id, lvl.level);
  //                 if (lvl.level > 0) {
  //                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
  //                         "counter-reset": counter
  //                     });
  //                 } else {
  //                     topCounters.push(counter);
  //                 }
  //                 css += this.styleToString(`p.${className}:before`, {
  //                     "content": this.levelTextToContent(lvl.text, num.id),
  //                     "counter-increment": counter
  //                 });
  //             } else if(lvl.bulletPictureId) {
  //                 let pict = bulletMap[lvl.bulletPictureId];
  //                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();
  //                 css += this.styleToString(`p.${className}:before`, {
  //                     "content": "' '",
  //                     "display": "inline-block",
  //                     "background": `var(${variable})`
  //                 }, pict.style);
  //                 this.document.loadNumberingImage(pict.referenceId).then(data => {
  //                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
  //                     container.appendChild(createStyleElement(text));
  //                 });
  //             } else {
  //                 listStyleType = this.numFormatToCssValue(lvl.format);
  //             }
  //             css += this.styleToString(`p.${className}`, {
  //                 "display": "list-item",
  //                 "list-style-position": "inside",
  //                 "list-style-type": listStyleType,
  //                 //TODO
  //                 //...num.style
  //             });
  //         }
  //     }
  //     if (topCounters.length > 0) {
  //         css += this.styleToString(`.${this.className}-wrapper`, {
  //             "counter-reset": topCounters.join(" ")
  //         });
  //     }
  //     return createStyleElement(css);
  // }
  renderNumbering(numberings, styleContainer) {
    var styleText = "";
    var resetCounters = [];
    for (var num of numberings) {
      var selector = `p.${this.numberingClass(num.id, num.level)}`;
      var listStyleType = "none";
      if (num.bullet) {
        let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
        styleText += this.styleToString(`${selector}:before`, {
          "content": "' '",
          "display": "inline-block",
          "background": `var(${valiable})`
        }, num.bullet.style);
        this.tasks.push(this.document.loadNumberingImage(num.bullet.src).then((data) => {
          var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
          styleContainer.appendChild(this.createStyleElement(text));
        }));
      } else if (num.levelText) {
        let counter = this.numberingCounter(num.id, num.level);
        const counterReset = counter + " " + (num.start - 1);
        if (num.level > 0) {
          styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
            "counter-set": counterReset
          });
        }
        resetCounters.push(counterReset);
        styleText += this.styleToString(`${selector}:before`, {
          "content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
          "counter-increment": counter,
          ...num.rStyle
        });
      } else {
        listStyleType = this.numFormatToCssValue(num.format);
      }
      styleText += this.styleToString(selector, {
        "display": "list-item",
        "list-style-position": "inside",
        "list-style-type": listStyleType,
        ...num.pStyle
      });
    }
    if (resetCounters.length > 0) {
      styleText += this.styleToString(this.rootSelector, {
        "counter-reset": resetCounters.join(" ")
      });
    }
    return this.createStyleElement(styleText);
  }
  renderStyles(styles) {
    var styleText = "";
    const stylesMap = this.styleMap;
    const defautStyles = keyBy(styles.filter((s) => s.isDefault), (s) => s.target);
    for (const style of styles) {
      var subStyles = style.styles;
      if (style.linked) {
        var linkedStyle = style.linked && stylesMap[style.linked];
        if (linkedStyle)
          subStyles = subStyles.concat(linkedStyle.styles);
        else if (this.options.debug)
          console.warn(`Can't find linked style ${style.linked}`);
      }
      for (const subStyle of subStyles) {
        var selector = `${style.target ?? ""}.${style.cssName}`;
        if (style.target != subStyle.target)
          selector += ` ${subStyle.target}`;
        if (defautStyles[style.target] == style)
          selector = `.${this.className} ${subStyle.target}, ` + selector;
        styleText += this.styleToString(selector, subStyle.values);
      }
    }
    return this.createStyleElement(styleText);
  }
  renderNotes(noteIds, notesMap, into) {
    var notes = noteIds.map((id) => notesMap[id]).filter((x) => x);
    if (notes.length > 0) {
      var result = this.createElement("ol", null, this.renderElements(notes));
      into.appendChild(result);
    }
  }
  renderElement(elem) {
    switch (elem.type) {
      case "paragraph" /* Paragraph */:
        return this.renderParagraph(elem);
      case "bookmarkStart" /* BookmarkStart */:
        return this.renderBookmarkStart(elem);
      case "bookmarkEnd" /* BookmarkEnd */:
        return null;
      //ignore bookmark end
      case "run" /* Run */:
        return this.renderRun(elem);
      case "table" /* Table */:
        return this.renderTable(elem);
      case "row" /* Row */:
        return this.renderTableRow(elem);
      case "cell" /* Cell */:
        return this.renderTableCell(elem);
      case "hyperlink" /* Hyperlink */:
        return this.renderHyperlink(elem);
      case "smartTag" /* SmartTag */:
        return this.renderSmartTag(elem);
      case "drawing" /* Drawing */:
        return this.renderDrawing(elem);
      case "image" /* Image */:
        return this.renderImage(elem);
      case "text" /* Text */:
        return this.renderText(elem);
      case "text" /* Text */:
        return this.renderText(elem);
      case "deletedText" /* DeletedText */:
        return this.renderDeletedText(elem);
      case "tab" /* Tab */:
        return this.renderTab(elem);
      case "symbol" /* Symbol */:
        return this.renderSymbol(elem);
      case "break" /* Break */:
        return this.renderBreak(elem);
      case "footer" /* Footer */:
        return this.renderContainer(elem, "footer");
      case "header" /* Header */:
        return this.renderContainer(elem, "header");
      case "footnote" /* Footnote */:
      case "endnote" /* Endnote */:
        return this.renderContainer(elem, "li");
      case "footnoteReference" /* FootnoteReference */:
        return this.renderFootnoteReference(elem);
      case "endnoteReference" /* EndnoteReference */:
        return this.renderEndnoteReference(elem);
      case "noBreakHyphen" /* NoBreakHyphen */:
        return this.createElement("wbr");
      case "vmlPicture" /* VmlPicture */:
        return this.renderVmlPicture(elem);
      case "vmlElement" /* VmlElement */:
        return this.renderVmlElement(elem);
      case "mmlMath" /* MmlMath */:
        return this.renderContainerNS(elem, ns2.mathML, "math", { xmlns: ns2.mathML });
      case "mmlMathParagraph" /* MmlMathParagraph */:
        return this.renderContainer(elem, "span");
      case "mmlFraction" /* MmlFraction */:
        return this.renderContainerNS(elem, ns2.mathML, "mfrac");
      case "mmlBase" /* MmlBase */:
        return this.renderContainerNS(
          elem,
          ns2.mathML,
          elem.parent.type == "mmlMatrixRow" /* MmlMatrixRow */ ? "mtd" : "mrow"
        );
      case "mmlNumerator" /* MmlNumerator */:
      case "mmlDenominator" /* MmlDenominator */:
      case "mmlFunction" /* MmlFunction */:
      case "mmlLimit" /* MmlLimit */:
      case "mmlBox" /* MmlBox */:
        return this.renderContainerNS(elem, ns2.mathML, "mrow");
      case "mmlGroupChar" /* MmlGroupChar */:
        return this.renderMmlGroupChar(elem);
      case "mmlLimitLower" /* MmlLimitLower */:
        return this.renderContainerNS(elem, ns2.mathML, "munder");
      case "mmlMatrix" /* MmlMatrix */:
        return this.renderContainerNS(elem, ns2.mathML, "mtable");
      case "mmlMatrixRow" /* MmlMatrixRow */:
        return this.renderContainerNS(elem, ns2.mathML, "mtr");
      case "mmlRadical" /* MmlRadical */:
        return this.renderMmlRadical(elem);
      case "mmlSuperscript" /* MmlSuperscript */:
        return this.renderContainerNS(elem, ns2.mathML, "msup");
      case "mmlSubscript" /* MmlSubscript */:
        return this.renderContainerNS(elem, ns2.mathML, "msub");
      case "mmlDegree" /* MmlDegree */:
      case "mmlSuperArgument" /* MmlSuperArgument */:
      case "mmlSubArgument" /* MmlSubArgument */:
        return this.renderContainerNS(elem, ns2.mathML, "mn");
      case "mmlFunctionName" /* MmlFunctionName */:
        return this.renderContainerNS(elem, ns2.mathML, "ms");
      case "mmlDelimiter" /* MmlDelimiter */:
        return this.renderMmlDelimiter(elem);
      case "mmlRun" /* MmlRun */:
        return this.renderMmlRun(elem);
      case "mmlNary" /* MmlNary */:
        return this.renderMmlNary(elem);
      case "mmlPreSubSuper" /* MmlPreSubSuper */:
        return this.renderMmlPreSubSuper(elem);
      case "mmlBar" /* MmlBar */:
        return this.renderMmlBar(elem);
      case "mmlEquationArray" /* MmlEquationArray */:
        return this.renderMllList(elem);
      case "inserted" /* Inserted */:
        return this.renderInserted(elem);
      case "deleted" /* Deleted */:
        return this.renderDeleted(elem);
      case "commentRangeStart" /* CommentRangeStart */:
        return this.renderCommentRangeStart(elem);
      case "commentRangeEnd" /* CommentRangeEnd */:
        return this.renderCommentRangeEnd(elem);
      case "commentReference" /* CommentReference */:
        return this.renderCommentReference(elem);
      case "altChunk" /* AltChunk */:
        return this.renderAltChunk(elem);
    }
    return null;
  }
  renderElements(elems, into) {
    if (elems == null)
      return null;
    var result = elems.flatMap((e) => this.renderElement(e)).filter((e) => e != null);
    if (into)
      appendChildren(into, result);
    return result;
  }
  renderContainer(elem, tagName, props) {
    return this.createElement(tagName, props, this.renderElements(elem.children));
  }
  renderContainerNS(elem, ns3, tagName, props) {
    return this.createElementNS(ns3, tagName, props, this.renderElements(elem.children));
  }
  renderParagraph(elem) {
    var result = this.renderContainer(elem, "p");
    const style = this.findStyle(elem.styleName);
    elem.tabs ?? (elem.tabs = style?.paragraphProps?.tabs);
    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    this.renderCommonProperties(result.style, elem);
    if (elem.runProps && elem.runProps.fontSize) {
      result.style.fontSize = elem.runProps.fontSize;
    }
    const numbering = elem.numbering ?? style?.paragraphProps?.numbering;
    if (numbering) {
      result.classList.add(this.numberingClass(numbering.id, numbering.level));
    }
    return result;
  }
  renderRunProperties(style, props) {
    this.renderCommonProperties(style, props);
  }
  renderCommonProperties(style, props) {
    if (props == null)
      return;
    if (props.color) {
      style["color"] = props.color;
    }
    if (props.fontSize) {
      style["font-size"] = props.fontSize;
    }
  }
  renderHyperlink(elem) {
    var result = this.renderContainer(elem, "a");
    this.renderStyleValues(elem.cssStyle, result);
    let href = "";
    if (elem.id) {
      const rel = this.document.documentPart.rels.find((it) => it.id == elem.id && it.targetMode === "External");
      href = rel?.target ?? href;
    }
    if (elem.anchor) {
      href += `#${elem.anchor}`;
    }
    result.href = href;
    return result;
  }
  renderSmartTag(elem) {
    return this.renderContainer(elem, "span");
  }
  renderCommentRangeStart(commentStart) {
    if (!this.options.renderComments)
      return null;
    const rng = new Range();
    this.commentHighlight?.add(rng);
    const result = this.createComment(`start of comment #${commentStart.id}`);
    this.later(() => rng.setStart(result, 0));
    this.commentMap[commentStart.id] = rng;
    return result;
  }
  renderCommentRangeEnd(commentEnd) {
    if (!this.options.renderComments)
      return null;
    const rng = this.commentMap[commentEnd.id];
    const result = this.createComment(`end of comment #${commentEnd.id}`);
    this.later(() => rng?.setEnd(result, 0));
    return result;
  }
  renderCommentReference(commentRef) {
    if (!this.options.renderComments)
      return null;
    var comment = this.document.commentsPart?.commentMap[commentRef.id];
    if (!comment)
      return null;
    const frg = new DocumentFragment();
    const commentRefEl = this.createElement("span", { className: `${this.className}-comment-ref` }, ["\u{1F4AC}"]);
    const commentsContainerEl = this.createElement("div", { className: `${this.className}-comment-popover` });
    this.renderCommentContent(comment, commentsContainerEl);
    frg.appendChild(this.createComment(`comment #${comment.id} by ${comment.author} on ${comment.date}`));
    frg.appendChild(commentRefEl);
    frg.appendChild(commentsContainerEl);
    return frg;
  }
  renderAltChunk(elem) {
    if (!this.options.renderAltChunks)
      return null;
    var result = this.createElement("iframe");
    this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then((x) => {
      result.srcdoc = x;
    }));
    return result;
  }
  renderCommentContent(comment, container) {
    container.appendChild(this.createElement("div", { className: `${this.className}-comment-author` }, [comment.author]));
    container.appendChild(this.createElement("div", { className: `${this.className}-comment-date` }, [new Date(comment.date).toLocaleString()]));
    this.renderElements(comment.children, container);
  }
  renderDrawing(elem) {
    var result = this.renderContainer(elem, "div");
    result.style.display = "inline-block";
    result.style.position = "relative";
    result.style.textIndent = "0px";
    this.renderStyleValues(elem.cssStyle, result);
    return result;
  }
  renderImage(elem) {
    let result = this.createElement("img");
    let transform = elem.cssStyle?.transform;
    this.renderStyleValues(elem.cssStyle, result);
    if (elem.srcRect && elem.srcRect.some((x) => x != 0)) {
      var [left, top, right, bottom] = elem.srcRect;
      transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
      result.style["clip-path"] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
    }
    if (elem.rotation)
      transform = `rotate(${elem.rotation}deg) ${transform ?? ""}`;
    result.style.transform = transform?.trim();
    if (this.document) {
      this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then((x) => {
        result.src = x;
      }));
    }
    return result;
  }
  renderText(elem) {
    return this.htmlDocument.createTextNode(elem.text);
  }
  renderDeletedText(elem) {
    return this.options.renderChanges ? this.renderText(elem) : null;
  }
  renderBreak(elem) {
    if (elem.break == "textWrapping") {
      return this.createElement("br");
    }
    return null;
  }
  renderInserted(elem) {
    if (this.options.renderChanges)
      return this.renderContainer(elem, "ins");
    return this.renderElements(elem.children);
  }
  renderDeleted(elem) {
    if (this.options.renderChanges)
      return this.renderContainer(elem, "del");
    return null;
  }
  renderSymbol(elem) {
    var span = this.createElement("span");
    span.style.fontFamily = elem.font;
    span.innerHTML = `&#x${elem.char};`;
    return span;
  }
  renderFootnoteReference(elem) {
    var result = this.createElement("sup");
    this.currentFootnoteIds.push(elem.id);
    result.textContent = `${this.currentFootnoteIds.length}`;
    return result;
  }
  renderEndnoteReference(elem) {
    var result = this.createElement("sup");
    this.currentEndnoteIds.push(elem.id);
    result.textContent = `${this.currentEndnoteIds.length}`;
    return result;
  }
  renderTab(elem) {
    var tabSpan = this.createElement("span");
    tabSpan.innerHTML = "&emsp;";
    if (this.options.experimental) {
      tabSpan.className = this.tabStopClass();
      var stops = findParent(elem, "paragraph" /* Paragraph */)?.tabs;
      this.currentTabs.push({ stops, span: tabSpan });
    }
    return tabSpan;
  }
  renderBookmarkStart(elem) {
    return this.createElement("span", { id: elem.name });
  }
  renderRun(elem) {
    if (elem.fieldRun)
      return null;
    const result = this.createElement("span");
    if (elem.id)
      result.id = elem.id;
    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    if (elem.verticalAlign) {
      const wrapper = this.createElement(elem.verticalAlign);
      this.renderElements(elem.children, wrapper);
      result.appendChild(wrapper);
    } else {
      this.renderElements(elem.children, result);
    }
    return result;
  }
  renderTable(elem) {
    let result = this.createElement("table");
    this.tableCellPositions.push(this.currentCellPosition);
    this.tableVerticalMerges.push(this.currentVerticalMerge);
    this.currentVerticalMerge = {};
    this.currentCellPosition = { col: 0, row: 0 };
    if (elem.columns)
      result.appendChild(this.renderTableColumns(elem.columns));
    this.renderClass(elem, result);
    this.renderElements(elem.children, result);
    this.renderStyleValues(elem.cssStyle, result);
    this.currentVerticalMerge = this.tableVerticalMerges.pop();
    this.currentCellPosition = this.tableCellPositions.pop();
    return result;
  }
  renderTableColumns(columns) {
    let result = this.createElement("colgroup");
    for (let col of columns) {
      let colElem = this.createElement("col");
      if (col.width)
        colElem.style.width = col.width;
      result.appendChild(colElem);
    }
    return result;
  }
  renderTableRow(elem) {
    let result = this.createElement("tr");
    this.currentCellPosition.col = 0;
    if (elem.gridBefore)
      result.appendChild(this.renderTableCellPlaceholder(elem.gridBefore));
    this.renderClass(elem, result);
    this.renderElements(elem.children, result);
    this.renderStyleValues(elem.cssStyle, result);
    if (elem.gridAfter)
      result.appendChild(this.renderTableCellPlaceholder(elem.gridAfter));
    this.currentCellPosition.row++;
    return result;
  }
  renderTableCellPlaceholder(colSpan) {
    const result = this.createElement("td", { colSpan });
    result.style["border"] = "none";
    return result;
  }
  renderTableCell(elem) {
    let result = this.renderContainer(elem, "td");
    const key = this.currentCellPosition.col;
    if (elem.verticalMerge) {
      if (elem.verticalMerge == "restart") {
        this.currentVerticalMerge[key] = result;
        result.rowSpan = 1;
      } else if (this.currentVerticalMerge[key]) {
        this.currentVerticalMerge[key].rowSpan += 1;
        result.style.display = "none";
      }
    } else {
      this.currentVerticalMerge[key] = null;
    }
    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    if (elem.span)
      result.colSpan = elem.span;
    this.currentCellPosition.col += result.colSpan;
    return result;
  }
  renderVmlPicture(elem) {
    return this.renderContainer(elem, "div");
  }
  renderVmlElement(elem) {
    var container = this.createSvgElement("svg");
    container.setAttribute("style", elem.cssStyle ? formatCssRules(elem.cssStyle) : "");
    const result = this.renderVmlChildElement(elem);
    if (elem.imageHref?.id) {
      this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart).then((x) => result.setAttribute("href", x)));
    }
    container.appendChild(result);
    requestAnimationFrame(() => {
      const bb = container.firstElementChild.getBBox();
      container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
      container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
    });
    return container;
  }
  renderVmlChildElement(elem) {
    const result = this.createSvgElement(elem.tagName);
    Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));
    for (let child of elem.children) {
      if (child.type == "vmlElement" /* VmlElement */) {
        result.appendChild(this.renderVmlChildElement(child));
      } else {
        result.appendChild(...asArray(this.renderElement(child)));
      }
    }
    return result;
  }
  renderMmlRadical(elem) {
    const base = elem.children.find((el) => el.type == "mmlBase" /* MmlBase */);
    if (elem.props?.hideDegree) {
      return this.createElementNS(ns2.mathML, "msqrt", null, this.renderElements([base]));
    }
    const degree = elem.children.find((el) => el.type == "mmlDegree" /* MmlDegree */);
    return this.createElementNS(ns2.mathML, "mroot", null, this.renderElements([base, degree]));
  }
  renderMmlDelimiter(elem) {
    const children = [];
    children.push(this.createElementNS(ns2.mathML, "mo", null, [elem.props.beginChar ?? "("]));
    children.push(...this.renderElements(elem.children));
    children.push(this.createElementNS(ns2.mathML, "mo", null, [elem.props.endChar ?? ")"]));
    return this.createElementNS(ns2.mathML, "mrow", null, children);
  }
  renderMmlNary(elem) {
    const children = [];
    const grouped = keyBy(elem.children, (x) => x.type);
    const sup = grouped["mmlSuperArgument" /* MmlSuperArgument */];
    const sub = grouped["mmlSubArgument" /* MmlSubArgument */];
    const supElem = sup ? this.createElementNS(ns2.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
    const subElem = sub ? this.createElementNS(ns2.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
    const charElem = this.createElementNS(ns2.mathML, "mo", null, [elem.props?.char ?? "\u222B"]);
    if (supElem || subElem) {
      children.push(this.createElementNS(ns2.mathML, "munderover", null, [charElem, subElem, supElem]));
    } else if (supElem) {
      children.push(this.createElementNS(ns2.mathML, "mover", null, [charElem, supElem]));
    } else if (subElem) {
      children.push(this.createElementNS(ns2.mathML, "munder", null, [charElem, subElem]));
    } else {
      children.push(charElem);
    }
    children.push(...this.renderElements(grouped["mmlBase" /* MmlBase */].children));
    return this.createElementNS(ns2.mathML, "mrow", null, children);
  }
  renderMmlPreSubSuper(elem) {
    const children = [];
    const grouped = keyBy(elem.children, (x) => x.type);
    const sup = grouped["mmlSuperArgument" /* MmlSuperArgument */];
    const sub = grouped["mmlSubArgument" /* MmlSubArgument */];
    const supElem = sup ? this.createElementNS(ns2.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
    const subElem = sub ? this.createElementNS(ns2.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
    const stubElem = this.createElementNS(ns2.mathML, "mo", null);
    children.push(this.createElementNS(ns2.mathML, "msubsup", null, [stubElem, subElem, supElem]));
    children.push(...this.renderElements(grouped["mmlBase" /* MmlBase */].children));
    return this.createElementNS(ns2.mathML, "mrow", null, children);
  }
  renderMmlGroupChar(elem) {
    const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
    const result = this.renderContainerNS(elem, ns2.mathML, tagName);
    if (elem.props.char) {
      result.appendChild(this.createElementNS(ns2.mathML, "mo", null, [elem.props.char]));
    }
    return result;
  }
  renderMmlBar(elem) {
    const result = this.renderContainerNS(elem, ns2.mathML, "mrow");
    switch (elem.props.position) {
      case "top":
        result.style.textDecoration = "overline";
        break;
      case "bottom":
        result.style.textDecoration = "underline";
        break;
    }
    return result;
  }
  renderMmlRun(elem) {
    const result = this.createElementNS(ns2.mathML, "ms", null, this.renderElements(elem.children));
    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    return result;
  }
  renderMllList(elem) {
    const result = this.createElementNS(ns2.mathML, "mtable");
    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    for (let child of this.renderElements(elem.children)) {
      result.appendChild(this.createElementNS(ns2.mathML, "mtr", null, [
        this.createElementNS(ns2.mathML, "mtd", null, [child])
      ]));
    }
    return result;
  }
  renderStyleValues(style, ouput) {
    for (let k in style) {
      if (k.startsWith("$")) {
        ouput.setAttribute(k.slice(1), style[k]);
      } else {
        ouput.style[k] = style[k];
      }
    }
  }
  renderClass(input, ouput) {
    if (input.className)
      ouput.className = input.className;
    if (input.styleName)
      ouput.classList.add(this.processStyleName(input.styleName));
  }
  findStyle(styleName) {
    return styleName && this.styleMap?.[styleName];
  }
  numberingClass(id, lvl) {
    return `${this.className}-num-${id}-${lvl}`;
  }
  tabStopClass() {
    return `${this.className}-tab-stop`;
  }
  styleToString(selectors, values2, cssText = null) {
    let result = `${selectors} {\r
`;
    for (const key in values2) {
      if (key.startsWith("$"))
        continue;
      result += `  ${key}: ${values2[key]};\r
`;
    }
    if (cssText)
      result += cssText;
    return result + "}\r\n";
  }
  numberingCounter(id, lvl) {
    return `${this.className}-num-${id}-${lvl}`;
  }
  levelTextToContent(text, suff, id, numformat) {
    const suffMap = {
      "tab": "\\9",
      "space": "\\a0"
    };
    var result = text.replace(/%\d*/g, (s) => {
      let lvl = parseInt(s.substring(1), 10) - 1;
      return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
    });
    return `"${result}${suffMap[suff] ?? ""}"`;
  }
  numFormatToCssValue(format) {
    return this.numFormatMapping[format] ?? format;
  }
  refreshTabStops() {
    if (!this.options.experimental)
      return;
    setTimeout(() => {
      const pixelToPoint = computePixelToPoint();
      for (let tab of this.currentTabs) {
        updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
      }
    }, 500);
  }
  createElementNS(ns3, tagName, props, children) {
    var result = ns3 ? this.htmlDocument.createElementNS(ns3, tagName) : this.htmlDocument.createElement(tagName);
    Object.assign(result, props);
    children && appendChildren(result, children);
    return result;
  }
  createElement(tagName, props, children) {
    return this.createElementNS(void 0, tagName, props, children);
  }
  createSvgElement(tagName, props, children) {
    return this.createElementNS(ns2.svg, tagName, props, children);
  }
  createStyleElement(cssText) {
    return this.createElement("style", { innerHTML: cssText });
  }
  createComment(text) {
    return this.htmlDocument.createComment(text);
  }
  later(func) {
    this.postRenderTasks.push(func);
  }
};
function removeAllElements(elem) {
  elem.innerHTML = "";
}
function appendChildren(elem, children) {
  children.forEach((c) => elem.appendChild(isString(c) ? document.createTextNode(c) : c));
}
function findParent(elem, type) {
  var parent = elem.parent;
  while (parent != null && parent.type != type)
    parent = parent.parent;
  return parent;
}

// src/docx-preview.ts
var defaultOptions = {
  ignoreHeight: false,
  ignoreWidth: false,
  ignoreFonts: false,
  breakPages: true,
  debug: false,
  experimental: false,
  className: "docx",
  inWrapper: true,
  hideWrapperOnPrint: false,
  trimXmlDeclaration: true,
  ignoreLastRenderedPageBreak: true,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  useBase64URL: false,
  renderChanges: false,
  renderComments: false,
  renderAltChunks: true
};
function parseAsync(data, userOptions) {
  const ops = { ...defaultOptions, ...userOptions };
  return WordDocument.load(data, new DocumentParser(ops), ops);
}
async function renderDocument(document2, bodyContainer, styleContainer, userOptions) {
  const ops = { ...defaultOptions, ...userOptions };
  const renderer = new HtmlRenderer(window.document);
  return await renderer.render(document2, bodyContainer, styleContainer, ops);
}
async function renderAsync(data, bodyContainer, styleContainer, userOptions) {
  const doc = await parseAsync(data, userOptions);
  await renderDocument(doc, bodyContainer, styleContainer, userOptions);
  return doc;
}
export {
  defaultOptions,
  parseAsync,
  renderAsync,
  renderDocument
};
//# sourceMappingURL=docx-preview.js.map