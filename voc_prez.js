const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, LevelFormat, PageNumber, Header, Footer,
  BorderStyle, ShadingType, PageBreak, ImageRun,
  Table, TableRow, TableCell, WidthType, VerticalAlign
} = require('/sessions/magical-great-gates/.npm-global/lib/node_modules/docx');
const fs = require('fs');

// Brand colours
const BRAND_NAVY   = "0a3457";
const BRAND_BLUE   = "324158";
const BRAND_GREEN  = "85b020";
const QUOTE_GREY   = "333d49";
const LIGHT_BG     = "F2F7E5";
const LIGHT_NAVY   = "EEF2F6";
const WHITE        = "FFFFFF";
const BLACK        = "000000";

const logoData = fs.readFileSync('/sessions/magical-great-gates/logo.png');

// --- Numbering -------------------------------------------------------------------
const numberingConfig = [
  {
    reference: "bullets",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 480, hanging: 280 }, spacing: { before: 60, after: 60 } } } }]
  }
];

// --- Core helpers ----------------------------------------------------------------

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function spacer(before = 120) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [new TextRun("")] });
}

function slideLabel(text) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 18, font: "Arial", color: BRAND_GREEN })]
  });
}

function slideTitle(text) {
  return new Paragraph({
    spacing: { before: 80, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND_GREEN, space: 6 } },
    children: [new TextRun({ text, bold: true, size: 40, font: "Arial", color: BRAND_NAVY })]
  });
}

function topicLabel(number, name) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [
      new TextRun({ text: `TOPIC ${number}  `, bold: true, size: 19, font: "Arial", color: BRAND_GREEN }),
      new TextRun({ text: name.toUpperCase(), bold: true, size: 19, font: "Arial", color: BRAND_BLUE })
    ]
  });
}

function insightText(text) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, size: 24, font: "Arial", color: BLACK })]
  });
}

// Primary quote with optional attribution
function prezQuote(text, attr) {
  const result = [];
  result.push(new Paragraph({
    spacing: { before: 140, after: attr ? 40 : 160 },
    indent: { left: 320, right: 320 },
    shading: { type: ShadingType.CLEAR, fill: LIGHT_BG },
    border: { left: { style: BorderStyle.SINGLE, size: 14, color: BRAND_GREEN, space: 6 } },
    children: [new TextRun({
      text: "\u201C" + text + "\u201D",
      italics: true, size: 23, font: "Arial", color: BRAND_NAVY
    })]
  }));
  if (attr) {
    result.push(new Paragraph({
      spacing: { before: 0, after: 160 },
      indent: { left: 320 },
      children: [new TextRun({
        text: "- " + attr,
        size: 17, font: "Arial", color: QUOTE_GREY, italics: false
      })]
    }));
  }
  return result;
}

// Sub-insight block: numbered heading + body text + supporting quotes (array of {text, attr})
function subInsight(number, headline, body, quotes) {
  const result = [];

  // Normalise quotes to array of {text, attr}
  const quotesArr = !quotes ? [] :
    Array.isArray(quotes) ? quotes :
    typeof quotes === 'string' ? [{ text: quotes, attr: null }] :
    [quotes];

  // Numbered headline bar
  result.push(new Paragraph({
    spacing: { before: 160, after: 40 },
    shading: { type: ShadingType.CLEAR, fill: LIGHT_NAVY },
    children: [
      new TextRun({ text: `  ${number}.  `, bold: true, size: 22, font: "Arial", color: BRAND_GREEN }),
      new TextRun({ text: headline, bold: true, size: 22, font: "Arial", color: BRAND_NAVY })
    ]
  }));

  // Body text (one size up)
  if (body) {
    result.push(new Paragraph({
      spacing: { before: 40, after: 40 },
      indent: { left: 280 },
      children: [new TextRun({ text: body, size: 22, font: "Arial", color: QUOTE_GREY })]
    }));
  }

  // Supporting quotes (up to 3)
  quotesArr.slice(0, 3).forEach(q => {
    const qtext = typeof q === 'string' ? q : q.text;
    const qattr = typeof q === 'string' ? null : q.attr;

    result.push(new Paragraph({
      spacing: { before: 40, after: qattr ? 20 : 40 },
      indent: { left: 280, right: 200 },
      shading: { type: ShadingType.CLEAR, fill: LIGHT_BG },
      children: [new TextRun({
        text: "\u201C" + qtext + "\u201D",
        italics: true, size: 21, font: "Arial", color: BRAND_BLUE
      })]
    }));
    if (qattr) {
      result.push(new Paragraph({
        spacing: { before: 0, after: 40 },
        indent: { left: 280 },
        children: [new TextRun({
          text: "- " + qattr,
          size: 17, font: "Arial", color: QUOTE_GREY
        })]
      }));
    }
  });

  return result;
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DDDDDD", space: 1 } },
    children: [new TextRun({ text: "DETAILS", bold: true, size: 17, font: "Arial", color: "AAAAAA" })]
  });
}

// Opportunity bullet
function oppBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 80, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: BLACK })]
  });
}

// Action bullet with bold lead-in
function actionBullet(lead, detail) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 80, after: 60 },
    children: [
      new TextRun({ text: lead + "  ", bold: true, size: 22, font: "Arial", color: BRAND_NAVY }),
      new TextRun({ text: detail, size: 22, font: "Arial", color: BLACK })
    ]
  });
}

function priorityBand(text) {
  return new Paragraph({
    spacing: { before: 280, after: 80 },
    shading: { type: ShadingType.CLEAR, fill: BRAND_NAVY },
    children: [new TextRun({ text: "  " + text.toUpperCase() + "  ", bold: true, size: 22, font: "Arial", color: WHITE })]
  });
}

// --- Table helpers ---------------------------------------------------------------
function npsCell(score, shade) {
  let colour = "000000";
  if (score === "10") colour = "5a7d10";
  else if (score === "9-10" || score === "~9-10") colour = "5a7d10";
  else if (score === "8" || score === "8-9" || score === "~8-9") colour = BRAND_BLUE;
  else if (score === "7") colour = "B8860B";
  else if (score === "6") colour = "C0392B";
  return new TableCell({
    shading: shade ? { type: ShadingType.CLEAR, fill: LIGHT_BG } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    width: { size: 700, type: WidthType.DXA },
    children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: score, bold: true, size: 20, font: "Arial", color: colour })]
    })]
  });
}
function hdrCell(text, w) {
  return new TableCell({ shading: { type: ShadingType.CLEAR, fill: BRAND_NAVY },
    verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 18, font: "Arial", color: WHITE })] })]
  });
}
function datCell(text, shade, w) {
  return new TableCell({ shading: shade ? { type: ShadingType.CLEAR, fill: LIGHT_BG } : undefined,
    verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text, size: 18, font: "Arial", color: BLACK })] })]
  });
}
function quoCell(text, shade, w) {
  return new TableCell({ shading: shade ? { type: ShadingType.CLEAR, fill: LIGHT_BG } : undefined,
    verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text: "\u201C" + text + "\u201D",
      italics: true, size: 17, font: "Arial", color: QUOTE_GREY })] })]
  });
}

// --- Slide builders --------------------------------------------------------------

// topicInsightSlide: core insight + primary quote (with attribution) + sub-insights
function topicInsightSlide(number, name, insight, coreQuote, coreAttr, subs) {
  const nodes = [
    topicLabel(number, name),
    slideTitle(name),
    slideLabel("Core Insight"),
    insightText(insight),
    ...prezQuote(coreQuote, coreAttr),
    divider()
  ];
  subs.forEach(s => nodes.push(...subInsight(s.n, s.headline, s.body, s.quotes)));
  nodes.push(pageBreak());
  return nodes;
}

function topicOppsSlide(number, name, conclusion, opps) {
  const nodes = [
    topicLabel(number, name),
    slideLabel("Opportunities"),
    new Paragraph({
      spacing: { before: 60, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BRAND_GREEN, space: 4 } },
      children: [new TextRun({ text: name, bold: true, size: 28, font: "Arial", color: BRAND_NAVY })]
    })
  ];
  if (conclusion) {
    nodes.push(new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: conclusion, size: 21, font: "Arial", color: QUOTE_GREY, italics: true })]
    }));
  }
  opps.forEach(o => nodes.push(oppBullet(o)));
  nodes.push(pageBreak());
  return nodes;
}

// --- Document content ------------------------------------------------------------
const children = [];

// ==============================================
// COVER
// ==============================================
const logoWidthEmu  = 2286000;
const logoHeightEmu = Math.round(logoWidthEmu * 284 / 800);

children.push(
  spacer(2400),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 320 },
    children: [new ImageRun({ type: "png", data: logoData,
      transformation: { width: Math.round(logoWidthEmu / 9144), height: Math.round(logoHeightEmu / 9144) },
      altText: { title: "2Flow logo", description: "2Flow company logo", name: "2FlowLogo" }
    })]
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BRAND_GREEN, space: 4 } },
    children: [new TextRun("")]
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 },
    children: [new TextRun({ text: "Voice of Customer", bold: true, size: 56, font: "Arial", color: BRAND_NAVY })]
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "Presentation Summary", bold: true, size: 40, font: "Arial", color: BRAND_BLUE })]
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 480 },
    children: [new TextRun({ text: "9 Customer Interviews  |  March 2026", size: 24, font: "Arial", color: QUOTE_GREY })]
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 60 },
    indent: { left: 720, right: 720 },
    shading: { type: ShadingType.CLEAR, fill: LIGHT_BG },
    border: { left: { style: BorderStyle.SINGLE, size: 16, color: BRAND_GREEN, space: 6 } },
    children: [new TextRun({ text: "\u201CYou\u2019re trusting a company with all of your stock, all of your growth.\u201D",
      italics: true, size: 30, font: "Arial", color: BRAND_NAVY, bold: true })]
  }),
  pageBreak()
);

// ==============================================
// WHO WE TALKED TO
// ==============================================
const CW = [1600, 1400, 700, 5226];
const custRows = [
  new TableRow({ tableHeader: true, children: [hdrCell("Customer",CW[0]),hdrCell("Company",CW[1]),hdrCell("Score",CW[2]),hdrCell("In their own words",CW[3])] }),
  new TableRow({ children: [datCell("Darren Grant",false,CW[0]),datCell("Probiotic.ie",false,CW[1]),npsCell("10",false),quoCell("Speed and customer service. Orders dispatch unbelievably fast.",false,CW[3])] }),
  new TableRow({ children: [datCell("Loren Purcell",true,CW[0]),datCell("3DPrinters.ie",true,CW[1]),npsCell("10",true),quoCell("They make my life very easy. Peace of mind. Confidence.",true,CW[3])] }),
  new TableRow({ children: [datCell("Mihael Sanko",false,CW[0]),datCell("Ancestral Cosmetics",false,CW[1]),npsCell("~9-10",false),quoCell("When I don\u2019t have to get involved a lot, that means everything is working.",false,CW[3])] }),
  new TableRow({ children: [datCell("Sinead O\u2019Donovan",true,CW[0]),datCell("Ormonde Jayne",true,CW[1]),npsCell("~9-10",true),quoCell("Trust is huge. You\u2019re trusting a company with all of your stock, all of your growth.",true,CW[3])] }),
  new TableRow({ children: [datCell("Cecilia Doyle",false,CW[0]),datCell("Coloplast",false,CW[1]),npsCell("8",false),quoCell("At least an eight. It would be a 10 if Mintsoft could link to our system.",false,CW[3])] }),
  new TableRow({ children: [datCell("Priya George",true,CW[0]),datCell("Unicity",true,CW[1]),npsCell("8-9",true),quoCell("Eight to nine, depending on size and what your business needs.",true,CW[3])] }),
  new TableRow({ children: [datCell("Emma Gibson",false,CW[0]),datCell("Skin Shop",false,CW[1]),npsCell("~8-9",false),quoCell("It\u2019s never once gone - oh my god, let\u2019s leave 2Flow. They take a huge weight off our plates.",false,CW[3])] }),
  new TableRow({ children: [datCell("Susanna O’Connor & Paula Geoghegan",true,CW[0]),datCell("Poco Beauty",true,CW[1]),npsCell("7",true),quoCell("We will fully admit we weren\u2019t prepared for how busy we would be. We have put them through a lot.",true,CW[3])] }),
  new TableRow({ children: [datCell("Danielle Goldbey",false,CW[0]),datCell("Spotlight Oral Care",false,CW[1]),npsCell("6",false),quoCell("I love the guys. But operationally it doesn\u2019t need to be as tense - and that\u2019s how I\u2019d describe the relationship over Christmas.",false,CW[3])] })
];

children.push(
  slideTitle("Who We Talked To"),
  new Paragraph({ spacing: { before: 40, after: 80 },
    children: [new TextRun({ text: "We spoke with 9 current customers. Across those who provided scores, the average Net Promoter Score was 8.4 out of 10 - a strong result that reflects genuine loyalty alongside clear areas for improvement.",
      size: 20, font: "Arial", color: BRAND_NAVY, bold: true })]
  }),
  new Paragraph({ spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: "Scores are on a 0-10 likelihood-to-recommend scale. Scores marked ~ are inferred from strong qualitative endorsements where no numeric score was given.",
      size: 19, font: "Arial", color: QUOTE_GREY, italics: true })]
  }),
  new Table({ width: { size: 8926, type: WidthType.DXA }, columnWidths: CW, rows: custRows }),
  pageBreak()
);

// ==============================================
// TOP INSIGHTS
// ==============================================
children.push(
  slideTitle("Top Insights"),
  new Paragraph({ spacing: { before: 40, after: 200 },
    children: [new TextRun({ text: "Six patterns that came through most clearly and most intensely - each pointing to a significant opportunity for the next phase of growth.",
      size: 22, font: "Arial", color: QUOTE_GREY, italics: true })]
  })
);

const insights = [
  { n:1, label:"Value", h:"Peace of mind is the product - and 2Flow barely says so",
    b:"Multiple customers reached for the same words without prompting: peace of mind, trust, confidence, weight off our plates. They are describing freedom from having to think about logistics at all. This is 2Flow\u2019s most powerful commercial message - and it is nearly invisible externally.",
    q:"When I don\u2019t have to get involved a lot, that means everything is working. That\u2019s what peace of mind means for me." },
  { n:2, label:"Operations", h:"Stock check-in unpredictability is the most widely felt daily frustration",
    b:"Four customers raised this independently. Stock arrives and goes silent - no timeline, no update. The frustration is not about speed; it is about the absence of any expectation to plan around. A written SLA and a proactive message when exceptions arise would remove most of the friction.",
    q:"Even if I know it will take a week, I can forecast for that. But sometimes it\u2019s really quick and sometimes it takes a really, really long time - and I just don\u2019t know." },
  { n:3, label:"Culture", h:"The \u2018say yes, make it happen\u2019 culture built this business - it needs to be the standard, not the exception",
    b:"The most consistent reason customers chose 2Flow - and stayed - is a bias toward yes. That culture is the single biggest commercial differentiator. The risk is that it lives in individuals rather than in the operation. When it slips, trust erodes fast and disproportionately.",
    q:"Others went off to review the logistics. He just said - yeah, give me the fridges. I hadn\u2019t come across that kind of flexibility before. That was it." },
  { n:4, label:"Peaks", h:"The highest-value clients had peak failures in 2024 and are heading into Q4 without a plan",
    b:"The customers who generate the most volume stretched 2Flow hardest last Black Friday - and came away with damaged confidence. Neither account has had a structured post-mortem or forward plan agreed. The window to act is before the summer.",
    q:"There was one particular weekend where I was checking on dispatch - no people dispatching. This is the week before Black Friday." },
  { n:5, label:"Systems", h:"Software friction is reducing satisfaction and capping what the relationship can be",
    b:"Software came up in six of nine interviews. The core issue is a Shopify-Mintsoft integration gap that forces every order amendment through a manual two-system process - a known limitation that competitors have already solved. One customer said her score would move from 8 to 10 if the systems could link.",
    q:"At least an eight. If their system could link to ours, it would be a 10." },
  { n:6, label:"Partner", h:"2Flow is a great reactive partner - the next phase requires becoming a proactive one",
    b:"Customers trust 2Flow to handle what is put in front of them. What several named directly is that 2Flow does not come to them with ideas, anticipate challenges, or develop new services ahead of demand. Competitors are already investing in TikTok Shop, Amazon prep, and account-level planning.",
    q:"Other fulfilment centres are setting up areas for TikTok Shop. They\u2019re actively going after this. We just don\u2019t see that on the 2Flow side." }
];

insights.forEach(i => {
  children.push(
    new Paragraph({ spacing: { before: 240, after: 40 },
      children: [
        new TextRun({ text: `${i.n}.  `, bold: true, size: 26, font: "Arial", color: BRAND_GREEN }),
        new TextRun({ text: `${i.label}  `, bold: true, size: 22, font: "Arial", color: BRAND_GREEN }),
        new TextRun({ text: i.h, bold: true, size: 26, font: "Arial", color: BRAND_NAVY })
      ]
    }),
    new Paragraph({ spacing: { before: 40, after: 80 },
      children: [new TextRun({ text: i.b, size: 21, font: "Arial", color: BLACK })]
    }),
    ...prezQuote(i.q, null)
  );
});
children.push(pageBreak());

// ==============================================
// TOPIC 1 - BEFORE 2FLOW
// ==============================================
children.push(...topicInsightSlide(1, "Before 2Flow",
  "Customers arrived from one of three broken situations: self-fulfilment that had buckled under growth, a previous 3PL that deflected blame while raising prices, or the wrong geography adding days and cost to every order. In most cases a specific trigger event forced the decision - not a gradual drift.",
  "I was packing the boxes myself. Backbreaking work. I would never go backwards to packing boxes again.",
  "Darren, Probiotic.ie",
  [
    { n:1, headline:"Self-fulfilment was a breaking point - growth forced the decision",
      body:"Businesses were packing orders themselves until growth made it untenable. Not a strategic choice - a survival one. COVID demand spikes, new retail channels, and physical limits all featured as triggers.",
      quotes:[
        { text:"COVID launched, catapulted my business... I realised I needed to scale, and I was looking for a partner.", attr:"Darren, Probiotic.ie" },
        { text:"We simply outgrew our ability to keep it in house, but we\u2019re not big enough to do our own warehouse.", attr:"Cecilia, Coloplast" }
      ]},
    { n:2, headline:"A previous 3PL was inefficient, kept raising prices, and deflected blame",
      body:"Several customers had already tried another provider. The pattern was consistent: operational delays, no accountability, and price increases that weren\u2019t justified by performance.",
      quotes:[
        { text:"Every Monday we had delays, and the weekend orders would spill into Tuesdays regularly. There were a lot of inefficiencies - and at the same time, the price kept going up.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"Even things like setting up a new product on the system was not a straightforward process.", attr:"Emma, Skin Shop" }
      ]},
    { n:3, headline:"Wrong geography was adding days, cost, and friction to every order",
      body:"Multiple customers were paying a daily penalty in time and money because their previous provider was in the wrong location for their market.",
      quotes:[
        { text:"We were in Galway, and most of our business was on this side of the country... We\u2019d add on an extra day or two, shipping stuff from Dublin to Galway for it to come back again.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"We were finding that there was quite a bit of a delay when we\u2019re sending it from Netherlands to Ireland, and the customers weren\u2019t quite happy with how long it was taking. As soon as we changed, it really made a difference.", attr:"Priya, Unicity" }
      ]},
    { n:4, headline:"A specific event triggered the move - it was rarely a gradual drift",
      body:"In almost every case, customers didn\u2019t leave their previous situation slowly. A growth event, a new channel, or a capacity crisis forced a hard decision on a short timeline.",
      quotes:[
        { text:"We simply outgrew our ability to keep it in house, but we\u2019re not big enough to do our own warehouse.", attr:"Cecilia, Coloplast" },
        { text:"At the time, we were only going into BT stores... it was a good time to move to them.", attr:"Susanna O’Connor, Poco Beauty" }
      ]}
  ]
));

children.push(...topicOppsSlide(1, "Before 2Flow",
  "Four actions coming directly from what customers told us about why they moved; which can improve the impact of our marketing and outbound sales activities.",
  [
    "Target the self-fulfilment inflection point with an explicit \u2018stop packing your own boxes\u2019 message in outbound sales and positioning.",
    "Win on accountability - make it a named brand differentiator. The number one complaint about previous providers was blame-shifting.",
    "Use the Dublin location advantage more actively. Multiple customers chose 2Flow partly because no credible Dublin-central alternative existed.",
    "Offer a low-friction onboarding package for sub-scale brands making the transition for the first time."
  ]
));

// ==============================================
// TOPIC 2 - HOW THEY FOUND 2FLOW
// ==============================================
children.push(...topicInsightSlide(2, "How They Found 2Flow",
  "Most customers found 2Flow via Google and made the decision based on the website and the first response - not after a lengthy pitch. Speed of reply and the feel of the first conversation did the selling. Strong advocates are already recommending 2Flow to peers, but this is entirely informal.",
  "I usually send an email out - who responds first? That\u2019s always a good sign for me. And I know John was pretty quick to respond.",
  "Mihael, Ancestral Cosmetics",
  [
    { n:1, headline:"Google was the primary discovery channel - the website was the credibility filter",
      body:"Two customers explicitly named Google as the starting point. One mentioned potentially asking ChatGPT first. In both cases, the website content was what converted the search into a conversation.",
      quotes:[
        { text:"The information on the website was so much information that to really give us an understanding already, straight off - okay, these are professionals.", attr:"Cecilia, Coloplast" },
        { text:"I probably just Googled them. I\u2019m not sure when it was - if it was ChatGPT. And then I went from Google and I saw they are real. They exist.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:2, headline:"Most ran a comparison - 2Flow won on responsiveness and human feel, not price",
      body:"Several customers interviewed multiple providers. What won was who responded fastest, felt most flexible, and gave a confident first impression.",
      quotes:[
        { text:"While others went off to review the logistics and everything else, John had just said, yeah, give me - send me fridges. I hadn\u2019t really come across a flexibility like that.", attr:"Darren, Probiotic.ie" },
        { text:"I usually send an email out - who responds first? That\u2019s always a good sign for me. And I know John was pretty quick to respond.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:3, headline:"Dublin location put 2Flow on the shortlist - geography filtered the competition",
      body:"For at least two customers, the Dublin-central location wasn\u2019t just a preference - it was the primary filter. 2Flow was the only credible option they found in that geography.",
      quotes:[
        { text:"2Flow was really the only one in that area. Then we just went to visit, met John, and I was like - okay, I have a good feeling about these guys.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"I think down to location.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:4, headline:"Word of mouth exists and is strong - but there is no structured referral channel",
      body:"There is no evidence of customers being directly referred via a formal programme. Several are strong advocates who recommend openly - this is not channelled or incentivised in any structured way.",
      quotes:[
        { text:"I would use a 2Flow-equivalent in every market.", attr:"Darren, Probiotic.ie" },
        { text:"I would definitely recommend them - if you want the easy-to-work warehouse that\u2019s efficient, does its job.", attr:"Mihael, Ancestral Cosmetics" }
      ]}
  ]
));

children.push(...topicOppsSlide(2, "How They Found 2Flow",
  "Four actions to strengthen the top of the funnel and activate existing advocates as a structured acquisition channel.",
  [
    "Invest in the website as the primary sales asset: sector pages for beauty, health, supplements; visible social proof; Trustpilot integration; high-intent SEO.",
    "Formalise a referral programme. Darren, Loren, Mihael, and Sinead O\u2019Donovan are ready to refer - none are on any structured programme.",
    "Set a first-response SLA for inbound leads (e.g. within 2 hours) and make it visible on the website.",
    "Build an AI-discovery presence: industry directory listings, authoritative reviews, and press mentions so 2Flow surfaces in AI-powered recommendations."
  ]
));

// ==============================================
// TOPIC 3 - WHY THEY CHOSE 2FLOW
// ==============================================
children.push(...topicInsightSlide(3, "Why They Chose 2Flow",
  "The decision came down to people and responsiveness, not price or capability. John won deals by doing one thing: saying yes quickly and meaning it. No hedging, no lengthy review, just commitment. John is retiring. The specific behaviours that won these customers need to be understood, documented, and deliberately transferred.",
  "Others went off to review the logistics. John just said - yeah, give me the fridges. I hadn\u2019t come across that kind of flexibility before. That was it.",
  "Darren, Probiotic.ie",
  [
    { n:1, headline:"John won the deal personally - and with a specific, repeatable pattern of behaviours",
      body:"Fast response, bias toward yes, willingness to commit before reviewing details, follow-through. Not just personality - a distinct set of transferable behaviours.",
      quotes:[
        { text:"I usually send an email out - who responds first? That\u2019s always a good sign for me. I went from Google and saw they are real, they exist. Then we just went to visit, met John, and I was like - okay, I have a good feeling about these guys.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"I interviewed 3 different places, and it was John who was dealing with me in 2Flow, and just liked them from the get-go. While others went off to review the logistics, John had just said, yeah, give me - send me fridges. I hadn\u2019t really come across a flexibility like that.", attr:"Darren, Probiotic.ie" }
      ]},
    { n:2, headline:"The \u2018yes first\u2019 mindset extended beyond the sale into ongoing account management",
      body:"John\u2019s responsiveness didn\u2019t stop at the initial deal. Customers reference proactive communication as an ongoing quality - not just a sales tactic.",
      quotes:[
        { text:"John had proactively told me we\u2019ve kind of maxed out your fridges, so we knew the capacity was being used.", attr:"Darren, Probiotic.ie" },
        { text:"John and Steven were always available. Whenever I had a question, they would respond quickly.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:3, headline:"The website created a professional first impression before any human contact",
      body:"Cecilia chose 2Flow before meeting anyone, based entirely on website content. The depth and quality of information was the credibility signal.",
      quotes:[
        { text:"We liked the information that was on their website. It felt like they would fit us quite well... straight off - okay, these are professionals.", attr:"Cecilia, Coloplast" }
      ]},
    { n:4, headline:"The in-person warehouse visit was consistently the tipping point",
      body:"Multiple customers referenced a physical visit as the moment the decision was made. Walking the floor and meeting the team gave confidence that no digital interaction could.",
      quotes:[
        { text:"We actually visited there once, and kind of went from there.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"They had very quick turnaround times of getting the orders out, and we would be able to do collections there.", attr:"Cecilia, Coloplast" }
      ]}
  ]
));

children.push(...topicOppsSlide(3, "Why They Chose 2Flow",
  "Four actions to systematise the sales behaviours that win customers - before the person who embodies them is no longer available.",
  [
    "Document and begin transferring John\u2019s sales behaviours before he leaves. Reply first, say yes before reviewing in detail, visit the prospect, follow up proactively. Write it down and work alongside him on live deals.",
    "Formalise the warehouse visit as a default step in every pitch process. A structured 30-minute tour as standard.",
    "Turn 2-3 customer stories into short website case studies. The fridge story, the first email response - these are what a prospect needs to read.",
    "Make the \u2018yes first\u2019 culture an explicit service standard in account management, not just in sales."
  ]
));

// ==============================================
// TOPIC 4 - CUSTOMER PROFILE & BUSINESS STAGE
// ==============================================
children.push(...topicInsightSlide(4, "Customer Profile & Business Stage",
  "Two distinct customer profiles are emerging: brands already scaling hard who feel every operational failure (the real ICP), and brands growing steadily but not yet under pressure (hopefuls). They need different things and carry different commercial risk. The distinction matters for account prioritisation, peak planning, and how 2Flow sells.",
  "I know my warehouse works. I know how things can be with two warehouses in the US, one in the UK. It doesn\u2019t need to be as tense - that\u2019s how I\u2019d describe the relationship over Christmas.",
  "Danielle, Spotlight Oral Care",
  [
    { n:1, headline:"DTC ecommerce is the dominant model - retail is a growing second channel",
      body:"The majority sell direct-to-consumer via Shopify. Retail (Boots, BT stores, independents, hospitals) is a growing secondary channel for several. Very few are pure retail.",
      quotes:[
        { text:"We\u2019ve always been selling directly via our website.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"I\u2019m in touch with the guys in the warehouse more than anyone else in the business... 2500 SKUs online.", attr:"Emma, Skin Shop" }
      ]},
    { n:2, headline:"Two distinct profiles: the real ICP is already under strain - and comparing alternatives",
      body:"The real ICP - Danielle, Susanna, Emma, Darren - is shipping at volume, has direct commercial pain when 2Flow underperforms, and is actively comparing options. The hopeful profile is growing but lower urgency and lower risk.",
      quotes:[
        { text:"We were snowed under for four weeks... the customer service was just like, I\u2019ve never seen it like that before.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"I know my warehouse works. I know how things can be with two warehouses in the US, one in the UK. It doesn\u2019t need to be as tense.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:3, headline:"The customer base is operationally lean - one or two people manage the entire relationship",
      body:"Across almost every interview, the 2Flow relationship is managed by one or two people stretched across multiple responsibilities. They have limited bandwidth for anything that requires active management.",
      quotes:[
        { text:"At the moment, I\u2019m doing both.", attr:"Sinead, Ormonde Jayne" },
        { text:"I\u2019m probably as involved with warehouse as you can get - as commercial manager.", attr:"Emma, Skin Shop" }
      ]},
    { n:4, headline:"Seasonality is real and universal - but the shape varies by business",
      body:"Every customer has peaks, but the triggers differ. Most peak in Q4. Some also spike on promotions, product launches, and gifting occasions. Shared forecasting is inconsistent.",
      quotes:[
        { text:"November, December, Black Friday, Christmas would be very busy in terms of the amount of orders.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"Q4 is usually big. And when we have a launch during the year, it might be a smaller peak.", attr:"Mihael, Ancestral Cosmetics" }
      ]}
  ]
));

children.push(...topicOppsSlide(4, "Customer Profile & Business Stage",
  "Four actions to prioritise the accounts most at risk and build service tracks suited to two distinct customer types.",
  [
    "Segment the client base by order volume, growth trajectory, and current friction level. High-volume, high-friction accounts are the retention priority.",
    "Build two distinct service tracks: proactive problem-solving and forward planning for the real ICP; responsive relationship-building for earlier-stage clients.",
    "Use the real ICP to define the sales target profile: strained-at-scale, Shopify-primary, Irish-market focus, enough order volume to make the relationship commercially meaningful.",
    "Position 2Flow as a service that reduces cognitive load for operationally lean teams. These customers do not have dedicated logistics managers."
  ]
));

// ==============================================
// TOPIC 5 - CORE VALUE
// ==============================================
children.push(...topicInsightSlide(5, "Core Value",
  "The product customers are actually buying is peace of mind - the ability to stop thinking about logistics and focus on running the business. They describe it in emotional terms: trust, confidence, relief, weight off our plates. This is not how 2Flow currently talks about itself. It should be.",
  "When I don\u2019t have to get involved a lot, that means everything is working. That\u2019s what peace of mind means for me.",
  "Mihael, Ancestral Cosmetics",
  [
    { n:1, headline:"Peace of mind is the dominant value - named directly and repeatedly",
      body:"Customers don\u2019t describe value in logistics terms. The commercial impact is that they can focus on growing the business rather than managing its operations.",
      quotes:[
        { text:"Trust is huge. You\u2019re trusting a company with all of your stock, all of your growth... trust is massive. And then obviously efficiency will be the big one as well.", attr:"Sinead, Ormonde Jayne" },
        { text:"Peace of mind. Confidence.", attr:"Loren, 3DPrinters.ie" },
        { text:"They take a huge weight off our plates.", attr:"Emma, Skin Shop" }
      ]},
    { n:2, headline:"Speed of dispatch directly drives customer satisfaction scores and revenue",
      body:"For at least two customers, 2Flow\u2019s dispatch speed is the primary commercial driver. It produces 5-star reviews, reduces customer service load, and directly supports brand reputation.",
      quotes:[
        { text:"The cutoff time is 1pm. A lot of the time the guys will still dispatch a 2pm order. The customer expected Wednesday - but the package arrives Tuesday, and they write a review.", attr:"Darren, Probiotic.ie" },
        { text:"Speed of service and customer service. Orders dispatch unbelievably fast. That results in 5-star Trustpilot reviews from me.", attr:"Darren, Probiotic.ie" }
      ]},
    { n:3, headline:"Responsiveness is the practical, daily expression of trust",
      body:"When things go wrong or when customers need something urgently, the speed and attitude of 2Flow\u2019s response is what maintains the relationship. Not SLAs - tone and willingness.",
      quotes:[
        { text:"I\u2019m never waiting any longer than an hour to get an answer back from them.", attr:"Sinead, Ormonde Jayne" },
        { text:"Whenever we ask - even the customer service - they\u2019re always nice, they\u2019re always polite, they\u2019re always fast to respond. So all that together gives us peace of mind.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:4, headline:"Freeing up founders to focus on growth is the hidden commercial value",
      body:"Multiple customers explicitly compared the cost of 2Flow against the value of their own time. The service unlocks attention that compounds into revenue.",
      quotes:[
        { text:"That takes away focus from far more important tasks when you\u2019re running an ecommerce business.", attr:"Emma, Skin Shop" }
      ]}
  ]
));

children.push(...topicOppsSlide(5, "Core Value",
  "Four actions to make 2Flow\u2019s strongest commercial message visible, measurable, and central to every sales conversation.",
  [
    "Lead with peace of mind in every commercial conversation - not throughput, SLAs, or storage rates. Proposals and the website should open with this.",
    "Make dispatch speed a named, measured brand promise: quantify it (\u2018X% of orders dispatched same day\u2019) and use it as social proof.",
    "Protect responsiveness as a non-negotiable service standard. Formalise it and communicate it explicitly to prospects and clients alike.",
    "Build and publish the \u2018growth enabler\u2019 case study: Darren ran a business from Spain on a laptop. That is the emotional outcome 2Flow delivers."
  ]
));

// ==============================================
// TOPIC 6 - HOW THE BUSINESS HAS CHANGED
// ==============================================
children.push(...topicInsightSlide(6, "How the Business Has Changed Since 2Flow",
  "Customers consistently report growth since joining 2Flow - often significant growth - but rarely attribute it directly to 2Flow. The more precise story is that 2Flow removed a constraint, and the business was then free to grow. The value is not what 2Flow built - it is what it unlocked.",
  "I moved to Spain in 2024 while I could run the business from a laptop.",
  "Darren, Probiotic.ie",
  [
    { n:1, headline:"Significant growth is common - and 2Flow absorbed most of it",
      body:"Multiple customers report strong growth since joining. The warehouse absorbed it without becoming the bottleneck. 2Flow was not the limiting factor - the business grew and the logistics held.",
      quotes:[
        { text:"We grew a lot since we started - probably more than double since we started with them. We grew in different markets.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"Q4 last year was pretty good, more than higher than our projections. So there were no issues.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:2, headline:"Operational freedom enabled remote working and leaner teams",
      body:"For some customers, 2Flow is what allows the business to run without a large operational headcount - or to run from another country entirely.",
      quotes:[
        { text:"I was packing the boxes myself when I started the company... I would never go backwards to packing boxes again.", attr:"Darren, Probiotic.ie" },
        { text:"If I have to be doing it, we\u2019d be getting a lot of complaints. Let\u2019s just say that much.", attr:"Loren, 3DPrinters.ie" }
      ]},
    { n:3, headline:"The move to retail was made possible by having a reliable fulfilment partner",
      body:"For Poco Beauty and Spotlight Oral Care, the move into retail happened concurrent with or shortly after joining 2Flow. A capable 3PL was a structural requirement for entering retail.",
      quotes:[
        { text:"There\u2019s a level of reassurance - we might mess something up - instead, because they\u2019ve been doing it for so long for us.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:4, headline:"Traceability and data improved - customers know more about their stock than before",
      body:"For Coloplast, the move to 2Flow brought operational intelligence for the first time: sample tracking, batch visibility, recall readiness. A meaningful differentiator for regulated industries.",
      quotes:[
        { text:"We are now actually able to properly track our sample orders. If we ever have a recall, we know that we will be able to trace where these have gone.", attr:"Cecilia, Coloplast" }
      ]}
  ]
));

children.push(...topicOppsSlide(6, "How the Business Has Changed Since 2Flow",
  "Three actions to turn customer growth stories into commercial proof points - and claim the growth narrative in sales and marketing.",
  [
    "Claim the growth narrative: \u2018we remove the constraint that holds your business back\u2019 - not \u2018we handle your logistics\u2019.",
    "Build case study content around the growth story - Darren\u2019s remote working, Mihael doubling in size, Cecilia\u2019s traceability gains.",
    "Position 2Flow as the partner for DTC brands entering retail for the first time. Poco Beauty and Spotlight Oral Care are the proof points."
  ]
));

// ==============================================
// TOPIC 7 - PEAK PERFORMANCE
// ==============================================
children.push(...topicInsightSlide(7, "Peak Performance",
  "Peak performance is the single biggest source of both trust and friction. When 2Flow absorbs a peak well, loyalty deepens. When it doesn\u2019t - two accounts can confirm it didn\u2019t in 2024 - the relationship comes under serious strain that takes months to recover. The variability is the problem.",
  "There was one particular weekend where I was checking on dispatch - no people dispatching. This is the week before Black Friday.",
  "Danielle, Spotlight Oral Care",
  [
    { n:1, headline:"When peaks go well, customers feel safe and become stronger advocates",
      body:"The customers whose peak experiences were positive describe genuine loyalty and elevated confidence. The service exceeding expectations - dispatching after the cutoff, delivering ahead of promise - turns customers into advocates.",
      quotes:[
        { text:"A lot of the time the guys will still dispatch a 2pm order after the cutoff. The customer expected Wednesday - but finds the package delivered on Tuesday, and writes a review.", attr:"Darren, Probiotic.ie" },
      ]},
    { n:2, headline:"When peaks break down, the damage is significant and lingers",
      body:"Two of 2Flow\u2019s highest-volume clients had material peak failures in 2024. Both were still processing the aftermath in March 2026. Neither had fully recovered confidence.",
      quotes:[
        { text:"Some of our online orders took up to two weeks to get to customers. We have a three to five day turnaround promise. We were snowed under for four weeks.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"This year was a little bit more disappointing... there was one particular weekend where I was checking on dispatch and we had no people dispatching. This is the week before Black Friday.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:3, headline:"The \u2018which channel do you want to prioritise?\u2019 question is commercially unacceptable",
      body:"Danielle\u2019s sharpest feedback was being asked to choose between retail and online during Black Friday. This question signals under-resourcing and places the problem back on the client.",
      quotes:[
        { text:"We were being asked - what do you want to prioritise, retail or online? For any business during that peak period, you can imagine it\u2019s not an ideal scenario.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:4, headline:"Weekend cover during peak is a structural gap customers have noticed",
      body:"Both Danielle and Poco referenced the absence of weekend staffing during their most critical trading days. Customers are now benchmarking this against competitors.",
      quotes:[
        { text:"Two of the guys had gone sick, so there was a delay. Maybe train a couple more people on brands so that you\u2019re not relying on only two people - especially in peak season.", attr:"Sinead, Ormonde Jayne" },
        { text:"I think it comes down to staff... separate teams for retail as opposed to online, so that both sides are being catered for equally.", attr:"Susanna O’Connor, Poco Beauty" }
      ]}
  ]
));

children.push(...topicOppsSlide(7, "Peak Performance",
  "Four actions to prevent peak failures from becoming churn - and to make the peak planning process a competitive differentiator.",
  [
    "Build a mandatory pre-peak briefing process for all active clients: 8 weeks before Black Friday - forecast review, staffing plan, channel priority agreement, SLA confirmation.",
    "Fix weekend cover permanently. It is now a competitive benchmark. Flexible staffing, a permanent weekend rota during Q4, or a contractual peak arrangement - and communicate it explicitly to clients.",
    "Never ask a client to choose between channels. That question is a symptom of under-resourcing, not a solution. The answer is a staffing plan.",
    "Communicate the peak planning process to prospects as a feature. It is a selling point, not just internal process."
  ]
));

// ==============================================
// TOPIC 8 - DEMAND PREDICTABILITY & STOCK
// ==============================================
children.push(...topicInsightSlide(8, "Demand Predictability & Stock Management",
  "The most consistent daily frustration across the customer base is the unpredictability of stock check-in timelines. Customers are not asking for faster - they are asking to know how long it will take so they can plan. The absence of any expectation forces manual workarounds, customer service escalations, and product availability problems.",
  "Even if I know my deliveries are going to take a week to check in, I can forecast for that. But sometimes it\u2019s really quick and sometimes it takes a really, really long time - and I just don\u2019t know.",
  "Emma, Skin Shop",
  [
    { n:1, headline:"Most customers share some forecast data - but the quality varies significantly",
      body:"Several make an effort to communicate upcoming spikes and launches. The information is often informal, incomplete, or late - and there is no structured process on either side.",
      quotes:[
        { text:"Probably not. I don\u2019t think we\u2019re very good on our side at maybe communicating when our sales are on or what\u2019s a big weekend.", attr:"Emma, Skin Shop" },
        { text:"I always try to share projections with them, especially if we are launching a new product, expecting more volume, if we are preparing for Black Friday or a sale.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"We did try and tell them how busy we were going to be, but I do think they slightly underestimated.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:2, headline:"Stock check-in timelines are unpredictable - and customers need transparency, not perfection",
      body:"Cash flow constraints mean stock often arrives close to peak, creating pressure precisely when the warehouse is at its busiest. The absence of a timeline forces website workarounds and customer service escalations.",
      quotes:[
        { text:"Cash flow means we can\u2019t afford to order six months in advance. So at the minute I\u2019m waiting for stock to be checked in that\u2019s been in the warehouse for a week and a half.", attr:"Emma, Skin Shop" },
        { text:"It only showed up on the system yesterday - eight days after the production order was placed.", attr:"Priya, Unicity" }
      ]},
    { n:3, headline:"Kitting and pre-build jobs are particularly hard to track",
      body:"Customers who use 2Flow for kitting find it impossible to track progress or get reliable completion timelines. The partial-batch upload process makes it nearly impossible to monitor in real time.",
      quotes:[
        { text:"We never could get them to tally what they said was done and what was done. After Christmas, 521 boxes were found - we always knew the quantities didn\u2019t match.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"Give me an expectation. Say, oh Emma - this week it\u2019s hectic. Could we just maybe break this down into stages?", attr:"Emma, Skin Shop" }
      ]},
    { n:4, headline:"ASN errors compound the problem when inbound stock is partially or incorrectly booked",
      body:"When inbound ASNs are not fully booked in - or partially booked with missing pieces unannounced - customers are left with inaccurate inventory and no visibility into the discrepancy.",
      quotes:[
        { text:"ASNs sometimes aren\u2019t booked in - or partially booked in - and you\u2019re not told if there are pieces missing.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"Please do not upload - check the discrepancy first. And what we get back is: stop, it\u2019s been uploaded. Not even an apology.", attr:"Danielle, Spotlight Oral Care" }
      ]}
  ]
));

children.push(...topicOppsSlide(8, "Demand Predictability & Stock Management",
  "Four actions to remove the most widely felt daily frustration across the current customer base - improving trust and reducing manual workarounds.",
  [
    "Publish and enforce a stock check-in SLA with proactive exception communication. Give clients a timeline. When exceptions arise, message them first - do not wait to be chased.",
    "Introduce a structured monthly forecasting template for clients: expected inbound, anticipated peak dates, kitting jobs in the queue. The act of asking forces forward planning on both sides.",
    "Fix kitting job visibility. Clients need a status view: queued, in progress, complete. Partial batch uploads create phantom inventory problems.",
    "Resolve the kitting completion upload process with Mintsoft - completed jobs should be uploaded as a single batch with a timestamp."
  ]
));

// ==============================================
// TOPIC 9 - WHO IS NOT A GOOD FIT
// ==============================================
children.push(...topicInsightSlide(9, "Who Is Not a Good Fit",
  "2Flow\u2019s sweet spot is lean, growing, DTC Shopify brands with manageable complexity. The gaps are high-volume pallet freight, cold chain wholesale, and large automated operations. The manual model works until it doesn\u2019t - and headcount dependency is a structural vulnerability at peak.",
  "From whatever I know, 2Flow is a lot more manual. If you had a bigger operation - like our European warehouse which handles hundreds of thousands of orders daily - you couldn\u2019t work like that.",
  "Priya, Unicity",
  [
    { n:1, headline:"The sweet spot is lean DTC brands with moderate complexity and manageable peaks",
      body:"Customers consistently describe 2Flow as ideal for pick-and-pack brands with a clear SKU range, Irish DTC or omnichannel, and manageable order volumes.",
      quotes:[
        { text:"Shopify stores, any store... a lot of people probably think they don\u2019t have enough SKUs, but I\u2019d probably recommend it to anyone selling physical goods.", attr:"Darren, Probiotic.ie" },
        { text:"If they\u2019re just ordering one SKU and it\u2019s going in a box and out the door - I can\u2019t imagine those companies would have any issues, unless it\u2019s volume.", attr:"Sinead, Ormonde Jayne" }
      ]},
    { n:2, headline:"The warehouse is largely manual - headcount dependency becomes a risk at peak",
      body:"The service is heavily dependent on headcount - and headcount is unpredictable during sickness, staff turnover, and peak periods. When people are not there, orders don\u2019t go out.",
      quotes:[
        { text:"There was one particular weekend where I was checking on dispatch and we had no people dispatching. This is the week before Black Friday.", attr:"Danielle, Spotlight Oral Care" },
        { text:"From whatever I know, 2Flow is a lot more manual. If you had a bigger operation - like our European warehouse which handles hundreds of thousands of orders daily - you couldn\u2019t work like that.", attr:"Priya, Unicity" }
      ]},
    { n:3, headline:"High-volume pallet freight, cold chain, and large wholesale are not the core model",
      body:"Mihael says 2Flow is not optimised for large wholesale freight. Darren\u2019s cold chain need is met at small scale, but his ability to grow wholesale is commercially constrained by the absence of proper cold chain infrastructure.",
      quotes:[
        { text:"My ability to grow the wholesale business is slightly limited by the fact that they don\u2019t have cold chain. I would send a hell of a lot more stock to 2Flow if they had cold chain capacity.", attr:"Darren, Probiotic.ie" },
        { text:"I don\u2019t think they\u2019re really optimised for big wholesale orders, or it would require a special setup on their side.", attr:"Mihael, Ancestral Cosmetics" }
      ]},
    { n:4, headline:"Ireland is the only English-speaking EU member - Dublin-based fulfilment is a commercially undermarketed ICP advantage",
      body:"Post-Brexit, Ireland is the only English-speaking EU member state. Brands using Ireland as an EU distribution hub - routing product through Dublin rather than the UK - are a precise ICP fit. Priya moved Unicity\u2019s EU routing from Netherlands to Dublin with immediate improvement. This is a positioning advantage 2Flow hasn\u2019t activated in sales or marketing.",
      quotes:[
        { text:"We were finding that there was quite a bit of a delay when we\u2019re sending it from Netherlands to Ireland, and the customers weren\u2019t quite happy with how long it was taking. As soon as we changed, it really made a difference.", attr:"Priya, Unicity" }
      ]}
  ]
));

children.push(...topicOppsSlide(9, "Who Is Not a Good Fit",
  "Five actions to sharpen the ideal customer profile, reduce structural operational risk, and activate 2Flow’s unique post-Brexit EU positioning advantage.",
  [
    "Use these boundaries to sharpen the ICP: DTC-primary, Shopify-based, Irish-market focus, 10-500 SKUs, manageable peaks. Tighten sales targeting accordingly.",
    "Invest in automation and cross-training to reduce human dependency. Document brand SOPs. Cross-train staff so no account depends on one or two specific people.",
    "Develop a formal peak staffing model: flexible agency agreements, standing weekend rotas in Q4, defined trigger points for additional resource.",
    "Make a deliberate decision on cold chain. One high-value customer would send significantly more stock if a proper cold chain wholesale service existed.",
    "Market 2Flow’s Dublin location as a post-Brexit EU distribution hub. Post-Brexit, Ireland is the only English-speaking EU member state. Target UK brands building EU presence and continental brands serving the Irish market — this is a positioning advantage that isn’t being used in sales or marketing."
  ]
));

// ==============================================
// TOPIC 10 - EXPANSION OPPORTUNITIES
// ==============================================
children.push(...topicInsightSlide(10, "Expansion Opportunities",
  "Four service extensions were named directly by customers - each addresses a real pain point that 2Flow is uniquely placed to solve. They are framed here as named products to be scoped, priced, and taken to market.",
  "Other fulfilment centres are setting up areas for TikTok Shop. They\u2019re actively going after this because they can see the potential there. We just don\u2019t see that on the 2Flow side.",
  "Danielle, Spotlight Oral Care",
  [
    { n:1, headline:"2Flow Delivery Support - outsourced tier-one customer service for order and delivery queries",
      body:"Multiple customers spend significant time fielding queries about order status and delivery timelines - information 2Flow already holds. 80% of Skin Shop\u2019s inbound customer service is order-related. The demand for outsourcing this is explicit.",
      quotes:[
        { text:"It\u2019s that kind of middleman - customer to us, to 2Flow, back to us, back to the customer. If it was something we outsourced, it would be a big help.", attr:"Emma, Skin Shop" },
        { text:"Yeah, yeah, provided they were able to handle the volume.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:2, headline:"2Flow Amazon Prep - structured FBA-ready fulfilment for brands on or entering Amazon",
      body:"Amazon is not a priority for most customers today, but it is in the thinking for several. Danielle is the most advanced - actively exploring Amazon.ie and pointing to Prime-level service as compelling.",
      quotes:[
        { text:"If they can offer 24-hour delivery like Prime, yeah, we\u2019d be in a really good position.", attr:"Danielle, Spotlight Oral Care" },
        { text:"We\u2019re starting to grow FBA in the UK too... trying to find alternative channels.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:3, headline:"2Flow TikTok Shop - warehouse infrastructure and fulfilment for social commerce",
      body:"TikTok Shop was raised as an area where competitor fulfilment centres are investing proactively. Emma tried to set it up and hit integration problems that were never resolved. Other providers are dedicating physical space and processes - and using it as a selling point.",
      quotes:[
        { text:"Other partners have sections in the warehouse, they have backdrops. They\u2019re actively going after this because they can see the potential there. And their knowledge of that is incredible.", attr:"Danielle, Spotlight Oral Care" },
        { text:"We set it up and there was such a huge issue with shipping that we actually stopped.", attr:"Emma, Skin Shop" }
      ]},
    { n:4, headline:"2Flow Kitting as a Service - structured, visible, scalable value-added processing",
      body:"Kitting, dekitting, personalisation, and pre-assembly are already being done, but informally - with no structured scoping, pricing, or visibility. The demand is there across multiple accounts.",
      quotes:[
        { text:"There was Christmas boxes flying out to everywhere. In hindsight, those boxes should have been pre-made before. We had 18,000 boxes to kit.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"We do have engraving. Works really well, we know, but there\u2019s a longer lead time on it, because it has to be taken out of the box, and the cap has to go to the engraving team and then put back in the box.", attr:"Sinead, Ormonde Jayne" }
      ]}
  ]
));

children.push(...topicOppsSlide(10, "Expansion Opportunities",
  "Four actions to capture incremental revenue from service extensions that customers have already asked for - reducing churn risk and creating new commercial relationships.",
  [
    "2Flow Delivery Support: pilot with Emma, Susanna, and Poco Beauty. Scope to tier-one order and delivery queries only, with a defined SLA and a modest monthly fee.",
    "2Flow Amazon Prep: build a clear FBA prep process - labelling, poly-bagging, bundling, FNSKU requirements. Danielle is the immediate prospect.",
    "2Flow TikTok Shop: dedicate a defined warehouse area, resolve the Shopify-TikTok shipping integration, and create a one-page proposition. Take it to the 3-4 clients most likely to benefit.",
    "2Flow Kitting as a Service: formal menu with lead times, per-unit pricing, and a job visibility tracker. Run structured kitting planning conversations before Q4."
  ]
));

// ==============================================
// TOPIC 11 - SYSTEMS & SOFTWARE
// ==============================================
children.push(...topicInsightSlide(11, "Systems & Software",
  "Software friction came up in six of nine interviews and traces back to three problems: Shopify-Mintsoft dual entry for order changes, no visibility into kitting and stock status, and customers never being shown how to use the system to its potential. These are all fixable.",
  "At least an eight. If their system could link to ours, it would be a 10.",
  "Cecilia, Coloplast",
  [
    { n:1, headline:"Shopify-Mintsoft dual entry is the most cited software frustration across the base",
      body:"Any order change requires manual updates in both systems. This is a known gap versus competitors. ShipHero was named as the benchmark - where integration is bi-directional and changes sync automatically.",
      quotes:[
        { text:"Whenever a customer wants to update the address, we just do it in Shopify and it syncs immediately. In Mintsoft, you need to change it in both systems.", attr:"Mihael, Ancestral Cosmetics" },
        { text:"Whenever a customer wants to update the order, we will have to alert 2Flow, they make the amendment on Mintsoft, and then we have to make a separate amendment on Shopify so that we reference it.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:2, headline:"Priority order flagging is missing - express orders get lost in the standard queue",
      body:"There is no mechanism to pull DPD express or next-day orders to the front of the dispatch queue. Back orders are also not being prioritised when stock arrives.",
      quotes:[
        { text:"There\u2019s a way on the system for them to flag the express deliveries and pull them first, because they are priority. If the customer has paid a premium to have them next day, they should go first.", attr:"Susanna O’Connor, Poco Beauty" },
        { text:"Those back orders, when they are released, they just get mixed into the orders that arrived that day. I would really like those orders to be at the top of the list.", attr:"Cecilia, Coloplast" }
      ]},
    { n:3, headline:"Inventory reporting and kitting tracking are not reliable enough to plan around",
      body:"The month-end inventory report cannot be trusted because it does not reflect stock sitting on the floor or kitting jobs that are only partially uploaded. Real-time reliability is absent.",
      quotes:[
        { text:"They send us an inventory at the end of every month, but it\u2019s not a true inventory if there\u2019s stock sitting on the floor that hasn\u2019t been booked in yet, or partially.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:4, headline:"Customers want Mintsoft training - and assume 2Flow would provide it if asked",
      body:"Several customers are not using Mintsoft\u2019s reporting capabilities - not because the capability isn\u2019t there, but because they have never been shown how. None had been proactively offered ongoing training.",
      quotes:[
        { text:"I probably would need a bit more training to know how to pull a report and get it really quickly. But I know that\u2019s available. I just probably have never approached them to do that.", attr:"Sinead, Ormonde Jayne" },
        { text:"We had training at the start. Before that, that was more for specific questions I had, but otherwise I was able to just use it for what I needed quite easily.", attr:"Priya, Unicity" }
      ]}
  ]
));

children.push(...topicOppsSlide(11, "Systems & Software",
  "Four actions to close the software gaps that are directly capping satisfaction scores - and to reduce the manual overhead that is frustrating the most active accounts.",
  [
    "Log with Mintsoft as priority: bi-directional Shopify order sync. Address changes made in Shopify should sync automatically. If Mintsoft cannot commit to a roadmap, evaluate ShipHero.",
    "Log with Mintsoft: priority order flagging in the dispatch queue. Express and next-day orders surface at the top. Back orders jump the queue when stock arrives.",
    "Introduce structured Mintsoft onboarding for all new clients: a 30-minute recorded session. And a 6-month check-in call for existing clients.",
    "Create a named manual integration SLA for clients whose systems cannot link to Mintsoft. Document the process, turnaround times, and responsibilities."
  ]
));

// ==============================================
// TOPIC 12 - ISSUES & FRICTION POINTS
// ==============================================
children.push(...topicInsightSlide(12, "Issues & Friction Points",
  "Friction is concentrated in five areas: stock check-in unpredictability, peak resourcing, defensive communication when problems arise, kitting job visibility, and software dual-entry. None individually fatal - but in combination they are eroding confidence in the accounts that matter most commercially.",
  "The thought process needs to change. It needs to be: hey, let us go and find a solution for you. Is that not what we get anymore?",
  "Danielle, Spotlight Oral Care",
  [
    { n:1, headline:"Priority 1 - Stock check-in delays and unpredictability (Emma, Susanna, Priya, Danielle)",
      body:"The most consistent daily frustration. Emma is forced to use a \u2018continue selling\u2019 workaround on her website because she cannot trust when stock will appear on-system. Priya waited 8 days for a production order.",
      quotes:[
        { text:"I don\u2019t know when or how long certain things are going to take. Even if I know it will take a week, then I can forecast for that.", attr:"Emma, Skin Shop" },
        { text:"Cash flow means we can’t afford to order six months in advance. So at the minute I’m waiting for stock to be checked in that’s been in the warehouse for a week and a half.", attr:"Emma, Skin Shop" }
      ]},
    { n:2, headline:"Priority 2 - Peak period resourcing and weekend cover (Danielle, Susanna)",
      body:"Two of the highest-volume clients had material peak failures in 2024 and are still in evaluation mode. No structured post-mortem has taken place. Without proactive intervention, the same failure is likely to repeat in Q4 2026.",
      quotes:[
        { text:"We were being asked - what do you want to prioritise, retail or online? For any business during that peak period, you can imagine it’s not an ideal scenario.", attr:"Danielle, Spotlight Oral Care" },
        { text:"I think it comes down to staff... separate teams for retail as opposed to online, so that both sides are being catered for equally.", attr:"Susanna O’Connor, Poco Beauty" }
      ]},
    { n:3, headline:"Priority 3 - Defensive communication when problems arise (Danielle, Susanna, Mihael)",
      body:"When issues are raised, the first instinct of some team members is to explain or justify rather than acknowledge and solve. Danielle has named this as a cultural issue, not an individual one.",
      quotes:[
        { text:"I just literally get a one-line back to say: stop, it\u2019s been uploaded. Not - apologies, should have flagged.", attr:"Danielle, Spotlight Oral Care" },
        { text:"The thought process needs to change. It needs to be: hey, let us go and find a solution for you. Is that not what we get anymore?", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:4, headline:"Priority 4 - Kitting job visibility and tracking (Susanna, Emma, Danielle)",
      body:"Once a kitting job is submitted, customers have no visibility into status. Jobs sit in the queue with no estimated completion time and no proactive update.",
      quotes:[
        { text:"There\u2019s a job I requested in December - it\u2019s only recently been done. Give me an expectation. Say, oh Emma - this week it\u2019s hectic. Could we just maybe break this down into stages?", attr:"Emma, Skin Shop" },
        { text:"There was Christmas boxes flying out to everywhere. In hindsight, those boxes should have been pre-made before. We had 18,000 boxes to kit.", attr:"Susanna O’Connor, Poco Beauty" }
      ]}
  ]
));

children.push(...topicOppsSlide(12, "Issues & Friction Points",
  "Four actions to address the specific issues currently eroding confidence in the accounts that matter most - before they become a churn event.",
  [
    "Priority 1 - Stock check-in SLA: publish a standard timeline. When exceptions arise, communicate proactively. Do not wait for the client to chase.",
    "Priority 2 - Peak resourcing: schedule post-mortem and forward planning sessions with Danielle and Susanna before May 2026. Agree a written peak plan with staffing commitments.",
    "Priority 3 - Communication culture: acknowledge-and-solve training for all customer-facing staff. First response to any problem is acknowledgement, not justification.",
    "Priority 4 - Kitting visibility: introduce a simple job tracker and log the batch-upload issue with Mintsoft as a formal feature request."
  ]
));

// ==============================================
// TOPIC 13 - SATISFACTION & ADVOCACY
// ==============================================
children.push(...topicInsightSlide(13, "Likelihood to Recommend & Advocacy",
  "The average likelihood-to-recommend score of 8.4/10 is strong for a B2B service business. The risk is concentration - the score is held up by exceptional advocates and held back by one account in active competitive evaluation. Four customers are ready to refer without prompting.",
  "How likely would you have been to recommend in September last year? Eight. And now? Six. Genuinely, kind of where we\u2019re at right now.",
  "Danielle, Spotlight Oral Care",
  [
    { n:1, headline:"The advocates are exceptional - and already selling 2Flow to others unprompted",
      body:"Several customers score 9-10 and actively recommend 2Flow without being asked. These are vocal advocates who bring 2Flow into peer conversations organically. None are on a formal referral programme.",
      quotes:[
        { text:"100%, 100%, there\u2019s no hesitation.", attr:"Loren, 3DPrinters.ie" },
        { text:"10. I\u2019m being generous, but I can\u2019t fault.", attr:"Darren, Probiotic.ie" },
        { text:"I\u2019d totally recommend. If I thought somebody was opening up a business and they need fulfilment, I\u2019d say absolutely go to 2Flow.", attr:"Sinead, Ormonde Jayne" }
      ]},
    { n:2, headline:"Mid-range recommendation scores are the majority - broadly positive but always with caveats",
      body:"Most customers would recommend 2Flow, but with caveats. The caveat is usually peak performance, software, or the need to understand the business\u2019s size and complexity.",
      quotes:[
        { text:"Eight to nine, depending on the size requirement and what your business needs.", attr:"Priya, Unicity" },
        { text:"At least an eight - would depend a little bit on what their needs were.", attr:"Cecilia, Coloplast" }
      ]},
    { n:3, headline:"Danielle\u2019s recommendation score dropped from 8 to 6 over 2024 - driven entirely by peak performance",
      body:"This is the clearest early-warning signal of churn risk in the customer base. Loyalty built over several years was eroded sharply by one peak failure. She remains warm personally but is conducting a formal annual competitive review.",
      quotes:[
        { text:"I have a great relationship with Simon, with John. But obviously, we have to look at this from an operational standpoint, and there are elements here that I\u2019m not so sure about.", attr:"Danielle, Spotlight Oral Care" },
        { text:"We\u2019ll weigh up: will it assist with the risk versus the consistency that we have. That\u2019s kind of what we do, like once a year.", attr:"Danielle, Spotlight Oral Care" }
      ]},
    { n:4, headline:"The language customers use when recommending is simple, consistent, and powerful",
      body:"When asked how they\u2019d describe 2Flow to a peer, customers use the same words: reliable, efficient, easy to work with, trustworthy, good people. This is 2Flow\u2019s brand promise as heard from its own customers.",
      quotes:[
        { text:"Reliable and efficient... trustworthy.", attr:"Sinead, Ormonde Jayne" },
        { text:"Easy to work with - efficient - does its job.", attr:"Mihael, Ancestral Cosmetics" },
      ]}
  ]
));

children.push(...topicOppsSlide(13, "Likelihood to Recommend & Advocacy",
  "Three actions to activate the advocacy that already exists - and prevent the most visible churn risk from materialising before Q4.",
  [
    "Activate the highest-scoring customers as a formal referral channel. Darren (10), Loren (10), Mihael and Sinead O\u2019Donovan (both ~9-10) are ready to refer. Create a named programme with a clear incentive.",
    "Intervene on Danielle\u2019s account before Q4 2026. A proactive meeting, a clear peak plan, and visible evidence that the 2024 feedback has been acted on is the minimum required.",
    "Use the customer\u2019s own language in marketing: \u2018Reliable. Efficient. Easy to work with. They just get it done.\u2019 These are not words from a brief - they are what actual customers say."
  ]
));

// ==============================================
// PRIORITY ACTIONS
// ==============================================
children.push(
  slideTitle("Priority Actions"),
  new Paragraph({ spacing: { before: 40, after: 100 },
    children: [new TextRun({ text: "The most important actions emerging from the VOC analysis, organised by time horizon.",
      size: 22, font: "Arial", color: QUOTE_GREY, italics: true })]
  }),

  priorityBand("Immediate - before summer 2026"),
  actionBullet("Meet Danielle and Susanna:", "Structured post-mortem of Q4 2024/Q4 2025, written peak plan for Q4 2026 with staffing commitments. Do this before May."),
  actionBullet("Publish a stock check-in SLA:", "Standard timeline, written. Proactive exception communication when it cannot be met."),
  actionBullet("Acknowledge-and-solve training:", "All customer-facing staff. First response to any problem is acknowledgement, not justification."),
  actionBullet("Launch a referral programme:", "Darren, Loren, Mihael, Sinead O\u2019Donovan are ready. Create a simple, named mechanism with a clear offer."),

  priorityBand("Short-term - Q3 2026"),
  actionBullet("Document John\u2019s sales behaviours:", "Before he leaves. Reply first, say yes, visit the prospect, follow up. Write it down and work alongside him on live deals."),
  actionBullet("Log with Mintsoft:", "Bi-directional Shopify sync and priority order flagging. If no roadmap, evaluate ShipHero as an alternative WMS."),
  actionBullet("Build a formal peak planning process:", "Mandatory pre-peak briefing for every client by end of July. Fix weekend cover permanently."),
  actionBullet("Pilot 2Flow Delivery Support:", "Outsourced tier-one delivery queries. Scope with Emma, Susanna, and Poco Beauty."),
  actionBullet("Productise kitting:", "Formal menu with lead times, per-unit pricing, and a simple job visibility tracker."),

  priorityBand("Strategic - build through 2026/2027"),
  actionBullet("Invest in the website:", "Sector pages, social proof, Trustpilot, ICP-targeted content, high-intent SEO. The website is the primary sales asset."),
  actionBullet("Segment the client base:", "By order volume, growth trajectory, and friction level. Build two distinct service tracks."),
  actionBullet("Scope Amazon Prep and TikTok Shop:", "Competitors are investing here now. Danielle is the immediate prospect for Amazon; Emma and Danielle for TikTok Shop."),
  actionBullet("Reduce headcount dependency:", "Automation investment, documented brand SOPs, cross-training programme. The manual model is a structural vulnerability at peak.")
);

// --- Build document ---------------------------------------------------------------
const doc = new Document({
  numbering: { config: numberingConfig },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [{
      id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 40, bold: true, font: "Arial", color: BRAND_NAVY },
      paragraph: { spacing: { before: 80, after: 160 }, outlineLevel: 0 }
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "2Flow Voice of Customer  |  March 2026  |  Page ", size: 17, font: "Arial", color: QUOTE_GREY }),
            new TextRun({ children: [PageNumber.CURRENT], size: 17, font: "Arial", color: QUOTE_GREY })
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/sessions/magical-great-gates/2Flow_VOC_Presentation.docx', buffer);
  console.log('Presentation copy written successfully.');
});
