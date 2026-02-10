const path = require('path');

const pptxgen = require('pptxgenjs');
const html2pptx = require('./.papert/skills/pptx/scripts/html2pptx');

async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  pptx.author = 'South Indian Bank';
  pptx.company = 'South Indian Bank';
  pptx.subject = 'Unaudited Standalone Financial Results';
  pptx.title = 'SIB Q3 FY 2025-26 Financial Results';

  const slidesDir = path.join(__dirname, 'workspace', 'slides');

  // Slide 1: Title
  await html2pptx(path.join(slidesDir, 'slide-01-title.html'), pptx);

  // Slide 2: Highlights
  await html2pptx(path.join(slidesDir, 'slide-02-highlights.html'), pptx);

  // Slide 3: Performance table
  {
    const { slide, placeholders } = await html2pptx(path.join(slidesDir, 'slide-03-performance-table.html'), pptx);
    const ph = placeholders.find(p => p.id === 'perfTable') || placeholders[0];

    const headerFill = { color: '0B1220' };
    const headerText = { color: 'FFFFFF', bold: true };

    const rows = [
      [
        { text: 'Particulars', options: { ...headerText, fill: headerFill } },
        { text: 'Q3 FY25-26\n(31 Dec 2025)', options: { ...headerText, fill: headerFill } },
        { text: 'Q2 FY25-26\n(30 Sep 2025)', options: { ...headerText, fill: headerFill } },
        { text: 'Q3 FY24-25\n(31 Dec 2024)', options: { ...headerText, fill: headerFill } }
      ],
      ['Total income', '3,00,346', '2,92,278', '2,77,996'],
      ['Interest expended', '1,63,685', '1,59,827', '1,50,148'],
      ['Operating profit', '58,433', '53,556', '52,884'],
      ['Provisions & contingencies', '8,041', '6,327', '6,604'],
      ['Profit before tax', '50,392', '47,229', '46,280'],
      ['Tax expense', '12,960', '12,092', '12,093'],
      ['Net profit', '37,432', '35,137', '34,187']
    ];

    slide.addTable(rows, {
      x: ph.x,
      y: ph.y,
      w: ph.w,
      h: ph.h,
      colW: [3.2, 2.0, 2.0, 2.0],
      border: { pt: 1, color: 'D9DEE8' },
      fontFace: 'Arial',
      fontSize: 11,
      valign: 'middle',
      align: 'center'
    });
  }

  // Slide 4: Segment results chart
  {
    const { slide, placeholders } = await html2pptx(path.join(slidesDir, 'slide-04-segment-results.html'), pptx);
    const ph = placeholders.find(p => p.id === 'segChart') || placeholders[0];

    slide.addChart(
      pptx.charts.BAR,
      [
        {
          name: 'PBT',
          labels: ['Treasury', 'Corporate/Wholesale', 'Retail', 'Other Ops'],
          values: [8138, 11518, 25683, 5053]
        }
      ],
      {
        ...ph,
        barDir: 'bar',
        showTitle: false,
        showLegend: false,
        showCatAxisTitle: true,
        catAxisTitle: 'Segment',
        showValAxisTitle: true,
        valAxisTitle: '₹ Lakhs',
        valAxisMinVal: 0,
        valAxisMajorUnit: 5000,
        dataLabelPosition: 'outEnd',
        dataLabelColor: '0B1220',
        chartColors: ['2D6CDF']
      }
    );
  }

  // Slide 5: Balance sheet table
  {
    const { slide, placeholders } = await html2pptx(path.join(slidesDir, 'slide-05-balance-sheet.html'), pptx);
    const ph = placeholders.find(p => p.id === 'bsTable') || placeholders[0];

    const headerFill = { color: '0B1220' };
    const headerText = { color: 'FFFFFF', bold: true };

    const rows = [
      [
        { text: 'Line item', options: { ...headerText, fill: headerFill } },
        { text: '31 Dec 2025', options: { ...headerText, fill: headerFill } },
        { text: '31 Mar 2025', options: { ...headerText, fill: headerFill } },
        { text: '31 Dec 2024', options: { ...headerText, fill: headerFill } }
      ],
      ['Deposits', '1,18,21,090', '1,07,52,560', '1,05,38,663'],
      ['Advances', '94,71,262', '85,68,207', '84,39,644'],
      ['Investments', '10,74,921', '9,83,829', '9,47,384'],
      ['Reserves & surplus', '10,74,921', '9,83,829', '9,47,384'],
      ['Total assets', '1,38,49,703', '1,24,65,512', '1,20,85,998']
    ];

    slide.addTable(rows, {
      x: ph.x,
      y: ph.y,
      w: ph.w,
      h: ph.h,
      colW: [3.0, 2.1, 2.1, 2.1],
      border: { pt: 1, color: 'D9DEE8' },
      fontFace: 'Arial',
      fontSize: 11,
      valign: 'middle',
      align: 'center'
    });
  }

  // Slide 6: Notes
  await html2pptx(path.join(slidesDir, 'slide-06-notes.html'), pptx);

  const out = path.join(__dirname, 'workspace', 'south-indian-bank-q3-fy25-26-results.pptx');
  await pptx.writeFile({ fileName: out });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
