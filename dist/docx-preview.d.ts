import JSZip from 'jszip';

declare enum DomType {
    Document = "document",
    Paragraph = "paragraph",
    Run = "run",
    Break = "break",
    NoBreakHyphen = "noBreakHyphen",
    Table = "table",
    Row = "row",
    Cell = "cell",
    Hyperlink = "hyperlink",
    SmartTag = "smartTag",
    Drawing = "drawing",
    Image = "image",
    Text = "text",
    Tab = "tab",
    Symbol = "symbol",
    BookmarkStart = "bookmarkStart",
    BookmarkEnd = "bookmarkEnd",
    Footer = "footer",
    Header = "header",
    FootnoteReference = "footnoteReference",
    EndnoteReference = "endnoteReference",
    Footnote = "footnote",
    Endnote = "endnote",
    SimpleField = "simpleField",
    ComplexField = "complexField",
    Instruction = "instruction",
    VmlPicture = "vmlPicture",
    MmlMath = "mmlMath",
    MmlMathParagraph = "mmlMathParagraph",
    MmlFraction = "mmlFraction",
    MmlFunction = "mmlFunction",
    MmlFunctionName = "mmlFunctionName",
    MmlNumerator = "mmlNumerator",
    MmlDenominator = "mmlDenominator",
    MmlRadical = "mmlRadical",
    MmlBase = "mmlBase",
    MmlDegree = "mmlDegree",
    MmlSuperscript = "mmlSuperscript",
    MmlSubscript = "mmlSubscript",
    MmlPreSubSuper = "mmlPreSubSuper",
    MmlSubArgument = "mmlSubArgument",
    MmlSuperArgument = "mmlSuperArgument",
    MmlNary = "mmlNary",
    MmlDelimiter = "mmlDelimiter",
    MmlRun = "mmlRun",
    MmlEquationArray = "mmlEquationArray",
    MmlLimit = "mmlLimit",
    MmlLimitLower = "mmlLimitLower",
    MmlMatrix = "mmlMatrix",
    MmlMatrixRow = "mmlMatrixRow",
    MmlBox = "mmlBox",
    MmlBar = "mmlBar",
    MmlGroupChar = "mmlGroupChar",
    VmlElement = "vmlElement",
    Inserted = "inserted",
    Deleted = "deleted",
    DeletedText = "deletedText",
    Comment = "comment",
    CommentReference = "commentReference",
    CommentRangeStart = "commentRangeStart",
    CommentRangeEnd = "commentRangeEnd",
    AltChunk = "altChunk"
}
interface OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    cssStyle?: Record<string, string>;
    props?: Record<string, any>;
    styleName?: string;
    className?: string;
    parent?: OpenXmlElement;
}
declare abstract class OpenXmlElementBase implements OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    cssStyle?: Record<string, string>;
    props?: Record<string, any>;
    className?: string;
    styleName?: string;
    parent?: OpenXmlElement;
}
interface WmlHyperlink extends OpenXmlElement {
    id?: string;
    anchor?: string;
}
interface WmlAltChunk extends OpenXmlElement {
    id?: string;
}
interface WmlSmartTag extends OpenXmlElement {
    uri?: string;
    element?: string;
}
interface WmlTable extends OpenXmlElement {
    columns?: WmlTableColumn[];
    cellStyle?: Record<string, string>;
    colBandSize?: number;
    rowBandSize?: number;
}
interface WmlTableRow extends OpenXmlElement {
    isHeader?: boolean;
    gridBefore?: number;
    gridAfter?: number;
}
interface WmlTableCell extends OpenXmlElement {
    verticalMerge?: 'restart' | 'continue' | string;
    span?: number;
}
interface IDomImage extends OpenXmlElement {
    src: string;
    srcRect: number[];
    rotation: number;
}
interface WmlTableColumn {
    width?: string;
}
interface IDomNumbering {
    id: string;
    level: number;
    start: number;
    pStyleName: string;
    pStyle: Record<string, string>;
    rStyle: Record<string, string>;
    levelText?: string;
    suff: string;
    format?: string;
    bullet?: NumberingPicBullet;
}
interface NumberingPicBullet {
    id: number;
    src: string;
    style?: string;
}

type LengthType = "px" | "pt" | "%" | "";
type Length = string;
interface CommonProperties {
    fontSize: Length;
    color: string;
}
type LengthUsageType = {
    mul: number;
    unit: LengthType;
    min?: number;
    max?: number;
};

declare class XmlParser {
    elements(elem: Element, localName?: string): Element[];
    element(elem: Element, localName: string): Element;
    elementAttr(elem: Element, localName: string, attrLocalName: string): string;
    attrs(elem: Element): Attr[];
    attr(elem: Element, localName: string): string;
    intAttr(node: Element, attrName: string, defaultValue?: number): number;
    hexAttr(node: Element, attrName: string, defaultValue?: number): number;
    floatAttr(node: Element, attrName: string, defaultValue?: number): number;
    boolAttr(node: Element, attrName: string, defaultValue?: boolean): boolean;
    lengthAttr(node: Element, attrName: string, usage?: LengthUsageType): Length;
}

interface Border {
    color: string;
    type: string;
    size: Length;
    frame: boolean;
    shadow: boolean;
    offset: Length;
}
interface Borders {
    top: Border;
    left: Border;
    right: Border;
    bottom: Border;
}

interface Column {
    space: Length;
    width: Length;
}
interface Columns {
    space: Length;
    numberOfColumns: number;
    separator: boolean;
    equalWidth: boolean;
    columns: Column[];
}
interface PageSize {
    width: Length;
    height: Length;
    orientation: "landscape" | string;
}
interface PageNumber {
    start: number;
    chapSep: "colon" | "emDash" | "endash" | "hyphen" | "period" | string;
    chapStyle: string;
    format: "none" | "cardinalText" | "decimal" | "decimalEnclosedCircle" | "decimalEnclosedFullstop" | "decimalEnclosedParen" | "decimalZero" | "lowerLetter" | "lowerRoman" | "ordinalText" | "upperLetter" | "upperRoman" | string;
}
interface PageMargins {
    top: Length;
    right: Length;
    bottom: Length;
    left: Length;
    header: Length;
    footer: Length;
    gutter: Length;
}
declare enum SectionType {
    Continuous = "continuous",
    NextPage = "nextPage",
    NextColumn = "nextColumn",
    EvenPage = "evenPage",
    OddPage = "oddPage"
}
interface FooterHeaderReference {
    id: string;
    type: string | "first" | "even" | "default";
}
interface SectionProperties {
    type: SectionType | string;
    pageSize: PageSize;
    pageMargins: PageMargins;
    pageBorders: Borders;
    pageNumber: PageNumber;
    columns: Columns;
    footerRefs: FooterHeaderReference[];
    headerRefs: FooterHeaderReference[];
    titlePage: boolean;
}

interface DocumentElement extends OpenXmlElement {
    props: SectionProperties;
}

interface LineSpacing {
    after: Length;
    before: Length;
    line: number;
    lineRule: "atLeast" | "exactly" | "auto";
}

interface WmlRun extends OpenXmlElement, RunProperties {
    id?: string;
    verticalAlign?: string;
    fieldRun?: boolean;
}
interface RunProperties extends CommonProperties {
}

interface WmlParagraph extends OpenXmlElement, ParagraphProperties {
}
interface ParagraphProperties extends CommonProperties {
    sectionProps: SectionProperties;
    tabs: ParagraphTab[];
    numbering: ParagraphNumbering;
    border: Borders;
    textAlignment: "auto" | "baseline" | "bottom" | "center" | "top" | string;
    lineSpacing: LineSpacing;
    keepLines: boolean;
    keepNext: boolean;
    pageBreakBefore: boolean;
    outlineLevel: number;
    styleName?: string;
    runProps: RunProperties;
}
interface ParagraphTab {
    style: "bar" | "center" | "clear" | "decimal" | "end" | "num" | "start" | "left" | "right";
    leader: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
    position: Length;
}
interface ParagraphNumbering {
    id: string;
    level: number;
}

interface IDomStyle {
    id: string;
    name?: string;
    cssName?: string;
    aliases?: string[];
    target: string;
    basedOn?: string;
    isDefault?: boolean;
    styles: IDomSubStyle[];
    linked?: string;
    next?: string;
    paragraphProps: ParagraphProperties;
    runProps: RunProperties;
}
interface IDomSubStyle {
    target: string;
    mod?: string;
    values: Record<string, string>;
}

interface DocumentParserOptions {
    ignoreWidth: boolean;
    debug: boolean;
}
declare class DocumentParser {
    options: DocumentParserOptions;
    constructor(options?: Partial<DocumentParserOptions>);
    parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[];
    parseComments(xmlDoc: Element): any[];
    parseDocumentFile(xmlDoc: Element): DocumentElement;
    parseBackground(elem: Element): any;
    parseBodyElements(element: Element): OpenXmlElement[];
    parseStylesFile(xstyles: Element): IDomStyle[];
    parseDefaultStyles(node: Element): IDomStyle;
    parseStyle(node: Element): IDomStyle;
    parseTableStyle(node: Element): IDomSubStyle[];
    parseNumberingFile(node: Element): IDomNumbering[];
    parseNumberingPicBullet(elem: Element): NumberingPicBullet;
    parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[];
    parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering;
    parseSdt(node: Element, parser: Function): OpenXmlElement[];
    parseInserted(node: Element, parentParser: Function): OpenXmlElement;
    parseDeleted(node: Element, parentParser: Function): OpenXmlElement;
    parseAltChunk(node: Element): WmlAltChunk;
    parseParagraph(node: Element): OpenXmlElement;
    parseParagraphProperties(elem: Element, paragraph: WmlParagraph): void;
    parseFrame(node: Element, paragraph: WmlParagraph): void;
    parseHyperlink(node: Element, parent?: OpenXmlElement): WmlHyperlink;
    parseSmartTag(node: Element, parent?: OpenXmlElement): WmlSmartTag;
    parseRun(node: Element, parent?: OpenXmlElement): WmlRun;
    parseMathElement(elem: Element): OpenXmlElement;
    parseMathProperies(elem: Element): Record<string, any>;
    parseRunProperties(elem: Element, run: WmlRun): void;
    parseVmlPicture(elem: Element): OpenXmlElement;
    checkAlternateContent(elem: Element): Element;
    parseDrawing(node: Element): OpenXmlElement;
    parseDrawingWrapper(node: Element): OpenXmlElement;
    parseGraphic(elem: Element): OpenXmlElement;
    parsePicture(elem: Element): IDomImage;
    parseTable(node: Element): WmlTable;
    parseTableColumns(node: Element): WmlTableColumn[];
    parseTableProperties(elem: Element, table: WmlTable): void;
    parseTablePosition(node: Element, table: WmlTable): void;
    parseTableRow(node: Element): WmlTableRow;
    parseTableRowProperties(elem: Element, row: WmlTableRow): void;
    parseTableCell(node: Element): OpenXmlElement;
    parseTableCellProperties(elem: Element, cell: WmlTableCell): void;
    parseTableCellVerticalText(elem: Element, cell: WmlTableCell): void;
    parseDefaultProperties(elem: Element, style?: Record<string, string>, childStyle?: Record<string, string>, handler?: (prop: Element) => boolean): Record<string, string>;
    parseUnderline(node: Element, style: Record<string, string>): void;
    parseFont(node: Element, style: Record<string, string>): void;
    parseIndentation(node: Element, style: Record<string, string>): void;
    parseSpacing(node: Element, style: Record<string, string>): void;
    parseMarginProperties(node: Element, output: Record<string, string>): void;
    parseTrHeight(node: Element, output: Record<string, string>): void;
    parseBorderProperties(node: Element, output: Record<string, string>): void;
}

interface Relationship {
    id: string;
    type: RelationshipTypes | string;
    target: string;
    targetMode: "" | "External" | string;
}
declare enum RelationshipTypes {
    OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    StylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
    Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
    Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    CustomProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties",
    Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    CommentsExtended = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    AltChunk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
}

interface OpenXmlPackageOptions {
    trimXmlDeclaration: boolean;
    keepOrigin: boolean;
}
declare class OpenXmlPackage {
    private _zip;
    options: OpenXmlPackageOptions;
    xmlParser: XmlParser;
    constructor(_zip: JSZip, options: OpenXmlPackageOptions);
    get(path: string): any;
    update(path: string, content: any): void;
    static load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage>;
    save(type?: any): Promise<any>;
    load(path: string, type?: JSZip.OutputType): Promise<any>;
    loadRelationships(path?: string): Promise<Relationship[]>;
    parseXmlDocument(txt: string): Document;
}

declare class Part {
    protected _package: OpenXmlPackage;
    path: string;
    protected _xmlDocument: Document;
    rels: Relationship[];
    constructor(_package: OpenXmlPackage, path: string);
    load(): Promise<any>;
    save(): void;
    protected parseXml(root: Element): void;
}

interface FontDeclaration {
    name: string;
    altName?: string;
    family: string;
    embedFontRefs: EmbedFontRef[];
}
interface EmbedFontRef {
    id: string;
    key: string;
    type: 'regular' | 'bold' | 'italic' | 'boldItalic';
}

declare class FontTablePart extends Part {
    fonts: FontDeclaration[];
    parseXml(root: Element): void;
}

declare class DocumentPart extends Part {
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    body: DocumentElement;
    parseXml(root: Element): void;
}

interface NumberingPartProperties {
    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
}
interface Numbering {
    id: string;
    abstractId: string;
    overrides: NumberingLevelOverride[];
}
interface NumberingLevelOverride {
    level: number;
    start: number;
    numberingLevel: NumberingLevel;
}
interface AbstractNumbering {
    id: string;
    name: string;
    multiLevelType: "singleLevel" | "multiLevel" | "hybridMultilevel" | string;
    levels: NumberingLevel[];
    numberingStyleLink: string;
    styleLink: string;
}
interface NumberingLevel {
    level: number;
    start: string;
    restart: number;
    format: 'lowerRoman' | 'lowerLetter' | string;
    text: string;
    justification: string;
    bulletPictureId: string;
    paragraphStyle: string;
    paragraphProps: ParagraphProperties;
    runProps: RunProperties;
}
interface NumberingBulletPicture {
    id: string;
    referenceId: string;
    style: string;
}

declare class NumberingPart extends Part implements NumberingPartProperties {
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
    domNumberings: IDomNumbering[];
    parseXml(root: Element): void;
}

declare class StylesPart extends Part {
    styles: IDomStyle[];
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}

interface ExtendedPropsDeclaration {
    template: string;
    totalTime: number;
    pages: number;
    words: number;
    characters: number;
    application: string;
    lines: number;
    paragraphs: number;
    company: string;
    appVersion: string;
}

declare class ExtendedPropsPart extends Part {
    props: ExtendedPropsDeclaration;
    parseXml(root: Element): void;
}

interface CorePropsDeclaration {
    title: string;
    description: string;
    subject: string;
    creator: string;
    keywords: string;
    language: string;
    lastModifiedBy: string;
    revision: number;
}

declare class CorePropsPart extends Part {
    props: CorePropsDeclaration;
    parseXml(root: Element): void;
}

declare class DmlTheme {
    colorScheme: DmlColorScheme;
    fontScheme: DmlFontScheme;
}
interface DmlColorScheme {
    name: string;
    colors: Record<string, string>;
}
interface DmlFontScheme {
    name: string;
    majorFont: DmlFormInfo;
    minorFont: DmlFormInfo;
}
interface DmlFormInfo {
    latinTypeface: string;
    eaTypeface: string;
    csTypeface: string;
}

declare class ThemePart extends Part {
    theme: DmlTheme;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}

declare abstract class WmlBaseNote implements OpenXmlElementBase {
    type: DomType;
    id: string;
    noteType: string;
}
declare class WmlFootnote extends WmlBaseNote {
    type: DomType;
}
declare class WmlEndnote extends WmlBaseNote {
    type: DomType;
}

declare class BaseNotePart<T extends WmlBaseNote> extends Part {
    protected _documentParser: DocumentParser;
    notes: T[];
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
}
declare class FootnotesPart extends BaseNotePart<WmlFootnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}
declare class EndnotesPart extends BaseNotePart<WmlEndnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}

interface WmlSettings {
    defaultTabStop: Length;
    footnoteProps: NoteProperties;
    endnoteProps: NoteProperties;
    autoHyphenation: boolean;
}
interface NoteProperties {
    nummeringFormat: string;
    defaultNoteIds: string[];
}

declare class SettingsPart extends Part {
    settings: WmlSettings;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}

declare class WmlComment extends OpenXmlElementBase {
    type: DomType;
    id: string;
    author: string;
    initials: string;
    date: string;
}

declare class CommentsPart extends Part {
    protected _documentParser: DocumentParser;
    comments: WmlComment[];
    commentMap: Record<string, WmlComment>;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}

type CommentsExtended = {
    paraId: string;
    paraIdParent?: string;
    done: boolean;
};
declare class CommentsExtendedPart extends Part {
    comments: CommentsExtended[];
    commentMap: Record<string, CommentsExtended>;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}

declare class WordDocument$1 {
    private _package;
    private _parser;
    private _options;
    rels: Relationship[];
    parts: Part[];
    partsMap: Record<string, Part>;
    documentPart: DocumentPart;
    fontTablePart: FontTablePart;
    numberingPart: NumberingPart;
    stylesPart: StylesPart;
    footnotesPart: FootnotesPart;
    endnotesPart: EndnotesPart;
    themePart: ThemePart;
    corePropsPart: CorePropsPart;
    extendedPropsPart: ExtendedPropsPart;
    settingsPart: SettingsPart;
    commentsPart: CommentsPart;
    commentsExtendedPart: CommentsExtendedPart;
    static load(blob: Blob | any, parser: DocumentParser, options: any): Promise<WordDocument$1>;
    save(type?: string): Promise<any>;
    private loadRelationshipPart;
    loadDocumentImage(id: string, part?: Part): Promise<string>;
    loadNumberingImage(id: string): Promise<string>;
    loadFont(id: string, key: string): Promise<string>;
    loadAltChunk(id: string, part?: Part): Promise<string>;
    private blobToURL;
    findPartByRelId(id: string, basePart?: Part): Part;
    getPathById(part: Part, id: string): string;
    private loadResource;
}

interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
}
declare const defaultOptions: Options;
type WordDocument = WordDocument$1;
declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument$1>;
declare function renderDocument(document: WordDocument, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<void>;
declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<WordDocument>;

export { type Options, type WordDocument, defaultOptions, parseAsync, renderAsync, renderDocument };
