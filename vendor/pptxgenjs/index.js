const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFileSync } = require("child_process");

const EMU_PER_INCH = 914400;
const DEFAULT_LAYOUT = { width: 13.333, height: 7.5 };

function toEmu(inches) {
  return Math.round(inches * EMU_PER_INCH);
}

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

class Slide {
  constructor() {
    this.shapes = [];
    this.background = null;
  }

  addText(text, options = {}) {
    this.shapes.push({ type: "text", text, options });
  }

  addShape(shapeType, options = {}) {
    this.shapes.push({ type: "shape", shapeType, options });
  }
}

class PptxGenJS {
  constructor() {
    this.layout = "LAYOUT_WIDE";
    this.author = "";
    this.company = "";
    this.subject = "";
    this.theme = {};
    this.slides = [];
  }

  addSlide() {
    const slide = new Slide();
    this.slides.push(slide);
    return slide;
  }

  writeFile({ fileName }) {
    return new Promise((resolve, reject) => {
      try {
        const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pptx-"));
        const parts = {
          root: tmpDir,
          rels: path.join(tmpDir, "_rels"),
          ppt: path.join(tmpDir, "ppt"),
          pptRels: path.join(tmpDir, "ppt", "_rels"),
          slides: path.join(tmpDir, "ppt", "slides"),
          slidesRels: path.join(tmpDir, "ppt", "slides", "_rels"),
          layouts: path.join(tmpDir, "ppt", "slideLayouts"),
          masters: path.join(tmpDir, "ppt", "slideMasters"),
          mastersRels: path.join(tmpDir, "ppt", "slideMasters", "_rels"),
          theme: path.join(tmpDir, "ppt", "theme"),
          docProps: path.join(tmpDir, "docProps")
        };

        Object.values(parts).forEach((dir) => {
          if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
          }
        });

        const slideCount = this.slides.length;
        const contentTypes = buildContentTypes(slideCount);
        fs.writeFileSync(path.join(tmpDir, "[Content_Types].xml"), contentTypes);

        fs.writeFileSync(path.join(parts.rels, ".rels"), buildRootRels());
        fs.writeFileSync(path.join(parts.docProps, "core.xml"), buildCoreProps(this));
        fs.writeFileSync(path.join(parts.docProps, "app.xml"), buildAppProps(slideCount));

        fs.writeFileSync(path.join(parts.ppt, "presentation.xml"), buildPresentationXml(slideCount));
        fs.writeFileSync(
          path.join(parts.pptRels, "presentation.xml.rels"),
          buildPresentationRels(slideCount)
        );

        fs.writeFileSync(path.join(parts.masters, "slideMaster1.xml"), buildSlideMasterXml());
        fs.writeFileSync(
          path.join(parts.mastersRels, "slideMaster1.xml.rels"),
          buildSlideMasterRels()
        );
        fs.writeFileSync(path.join(parts.layouts, "slideLayout1.xml"), buildSlideLayoutXml());
        fs.writeFileSync(path.join(parts.theme, "theme1.xml"), buildThemeXml(this.theme));

        this.slides.forEach((slide, index) => {
          const slideNumber = index + 1;
          fs.writeFileSync(
            path.join(parts.slides, `slide${slideNumber}.xml`),
            buildSlideXml(slide, slideNumber)
          );
          fs.writeFileSync(
            path.join(parts.slidesRels, `slide${slideNumber}.xml.rels`),
            buildSlideRels()
          );
        });

        const outDir = path.dirname(fileName);
        if (!fs.existsSync(outDir)) {
          fs.mkdirSync(outDir, { recursive: true });
        }

        execFileSync("zip", ["-X", "-r", fileName, "."], { cwd: tmpDir });
        resolve(fileName);
      } catch (error) {
        reject(error);
      }
    });
  }
}

PptxGenJS.prototype.ShapeType = {
  line: "line"
};

function buildContentTypes(slideCount) {
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    const num = index + 1;
    return `  <Override PartName=\"/ppt/slides/slide${num}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>`;
  }).join("\n");

  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation+xml\"/>
  <Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>
  <Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>
  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>
  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>
  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>
${slideOverrides}
</Types>`;
}

function buildRootRels() {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>
  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>
</Relationships>`;
}

function buildCoreProps(pptx) {
  const now = new Date().toISOString();
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">
  <dc:title>${xmlEscape(pptx.subject || "ChatGPT Deep Dive")}</dc:title>
  <dc:creator>${xmlEscape(pptx.author || "AntonAI")}</dc:creator>
  <cp:lastModifiedBy>${xmlEscape(pptx.author || "AntonAI")}</cp:lastModifiedBy>
  <dcterms:created xsi:type=\"dcterms:W3CDTF\">${now}</dcterms:created>
  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function buildAppProps(slideCount) {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">
  <Application>AntonAI Generator</Application>
  <Slides>${slideCount}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <PresentationFormat>Widescreen</PresentationFormat>
  <Company>AntonAI</Company>
  <AppVersion>1.0</AppVersion>
</Properties>`;
}

function buildPresentationXml(slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, index) => {
    const id = 256 + index;
    const rId = index + 2;
    return `    <p:sldId id=\"${id}\" r:id=\"rId${rId}\"/>`;
  }).join("\n");

  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:sldMasterIdLst>
    <p:sldMasterId id=\"1\" r:id=\"rId1\"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
${slideIds}
  </p:sldIdLst>
  <p:slideSize cx=\"12192000\" cy=\"6858000\" type=\"screen16x9\"/>
  <p:notesSz cx=\"6858000\" cy=\"9144000\"/>
</p:presentation>`;
}

function buildPresentationRels(slideCount) {
  const slideRels = Array.from({ length: slideCount }, (_, index) => {
    const num = index + 1;
    const rId = index + 2;
    return `  <Relationship Id=\"rId${rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide${num}.xml\"/>`;
  }).join("\n");

  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>
${slideRels}
</Relationships>`;
}

function buildSlideMasterXml() {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill>
          <a:srgbClr val=\"FFFFFF\"/>
        </a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id=\"1\" r:id=\"rId1\"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;
}

function buildSlideMasterRels() {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>
</Relationships>`;
}

function buildSlideLayoutXml() {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" type=\"blank\" preserve=\"1\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`;
}

function buildThemeXml(theme) {
  const font = theme.bodyFontFace || "Arial";
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"AntonAI\">
  <a:themeElements>
    <a:clrScheme name=\"AntonAI\">
      <a:dk1><a:srgbClr val=\"000000\"/></a:dk1>
      <a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>
      <a:dk2><a:srgbClr val=\"1F1F1F\"/></a:dk2>
      <a:lt2><a:srgbClr val=\"F2F2F2\"/></a:lt2>
      <a:accent1><a:srgbClr val=\"2B6CB0\"/></a:accent1>
      <a:accent2><a:srgbClr val=\"6366F1\"/></a:accent2>
      <a:accent3><a:srgbClr val=\"10B981\"/></a:accent3>
      <a:accent4><a:srgbClr val=\"F59E0B\"/></a:accent4>
      <a:accent5><a:srgbClr val=\"EF4444\"/></a:accent5>
      <a:accent6><a:srgbClr val=\"8B5CF6\"/></a:accent6>
      <a:hlink><a:srgbClr val=\"2B6CB0\"/></a:hlink>
      <a:folHlink><a:srgbClr val=\"1D4ED8\"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name=\"AntonAI\">
      <a:majorFont>
        <a:latin typeface=\"${xmlEscape(font)}\"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface=\"${xmlEscape(font)}\"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name=\"AntonAI\">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>
        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>
        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function buildSlideXml(slide, slideNumber) {
  const shapes = buildShapes(slide);
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
${shapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
}

function buildShapes(slide) {
  let shapeId = 2;
  return slide.shapes
    .map((shape) => {
      if (shape.type === "text") {
        const id = shapeId++;
        return buildTextShape(shape, id);
      }
      if (shape.type === "shape") {
        const id = shapeId++;
        return buildLineShape(shape, id);
      }
      return "";
    })
    .join("\n");
}

function buildTextShape(shape, id) {
  const opts = shape.options || {};
  const x = toEmu(opts.x || 0);
  const y = toEmu(opts.y || 0);
  const w = toEmu(opts.w || 1);
  const h = toEmu(opts.h || 1);
  const fontSize = Math.round((opts.fontSize || 18) * 100);
  const color = opts.color || "000000";
  const bold = opts.bold ? "1" : "0";
  const lines = String(shape.text).split("\n");
  const paragraphs = lines
    .map((line) => {
      return `      <a:p><a:r><a:rPr lang=\"ko-KR\" sz=\"${fontSize}\" b=\"${bold}\"><a:solidFill><a:srgbClr val=\"${color}\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`;
    })
    .join("\n");

  return `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id=\"${id}\" name=\"Text ${id}\"/>
          <p:cNvSpPr txBox=\"1\"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x=\"${x}\" y=\"${y}\"/>
            <a:ext cx=\"${w}\" cy=\"${h}\"/>
          </a:xfrm>
          <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap=\"square\"/>
          <a:lstStyle/>
${paragraphs}
        </p:txBody>
      </p:sp>`;
}

function buildLineShape(shape, id) {
  const opts = shape.options || {};
  const x = toEmu(opts.x || 0);
  const y = toEmu(opts.y || 0);
  const w = toEmu(opts.w || 1);
  const h = opts.h && opts.h > 0 ? toEmu(opts.h) : toEmu(0.02);
  const color = (opts.line && opts.line.color) || "2B6CB0";

  return `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id=\"${id}\" name=\"Line ${id}\"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x=\"${x}\" y=\"${y}\"/>
            <a:ext cx=\"${w}\" cy=\"${h}\"/>
          </a:xfrm>
          <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val=\"${color}\"/></a:solidFill>
          <a:ln w=\"0\"><a:noFill/></a:ln>
        </p:spPr>
      </p:sp>`;
}

function buildSlideRels() {
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>
</Relationships>`;
}

module.exports = PptxGenJS;
