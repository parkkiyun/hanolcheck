import { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, HeightRule, VerticalAlign } from 'docx';

const SINGLE_BORDER = { style: BorderStyle.SINGLE, size: 1, color: '000000' };

interface ReportData {
  type: '질병' | '출석인정' | '기타';
  schoolYear: string;
  year: string;
  class: string;
  number: string;
  name: string;
  reason: string;
  startDate: string;
  endDate: string;
  days: number;
  submissionDate: string;
  teacherName: string;
  guardianName?: string;
}

const createFooterNoteParagraphs = (type: ReportData['type']) => {
  if (type === '질병') {
    return {
      title: '질병 결석에 해당 하는 경우',
      paragraphs: [
        '1. 결석한 날부터 5일 이내에 의사의 진단서 또는 의견서(의사 소견서, 진료 확인서 등으로 병명, 진료기간 등이 기록된 증빙서류)를 첨부한 결석계를 제출하여 학교장의 승인을 받은 경우',
        '2. 다만, 상습적이지 않은 2일 이내의 결석은 질병으로 인한 결석임을 증명할 수 있는 자료(학부모 의견서, 처방전, 담임교사 확인서 등)가 첨부된 결석계를 5일 이내에 제출하여 학교장의 승인을 받은 경우',
      ],
    };
  }

  if (type === '기타') {
    return {
      title: '기타 결석에 해당 하는 경우',
      paragraphs: [
        '1. 기타 부득이한 사유로 학교장의 허가를 받은 경우',
      ],
    };
  }

  return {
    title: '출석 인정 결석에 해당 하는 경우',
    paragraphs: [
      '1. 천재지변 또는 법정 감염병, 공적의무(병역관계 등) 또는 공권력의 행사',
      '2. 학교장의 허가를 받은 ‘학교ㆍ시도(교육청)ㆍ국가를 대표한 대회 및 훈련 참가, 산업체 실습과정(현장실습, 현장실습과 연계한 취업), 교환학습, 교외체험학습, 「학교보건법」 제8조에 따른 등교중지’ 등으로 출석하지 못한 경우',
      '3. 「초ㆍ중등교육법 시행령」 제31조제1항에 따른 학교 내의 봉사, 사회봉사, 특별교육이수 기간',
      '4. 「초ㆍ중등교육법」제28조제6항에 따른 상담, 진로 프로그램 등 숙려제 참여 인정 기간',
      '5. 경조사',
      '6. 기타 부득이한 사유로 학교장의 허가를 받은 경우',
      '7.「학교폭력예방 및 대책에 관한 법률」 제12조에 따른 학교폭력대책심의위원회의 개최 및 동 위원회의 학교폭력 피해학생에 대한 보호조치 요청 이전에, 학교폭력 피해자가 학교폭력으로 인한 피해로 출석하지 못하였음을 같은 법 제14조제3항에 따른 학교폭력 전담기구의 조사 및 확인을 거쳐 학교의 장이 인정한 경우',
      '8. 경찰청 「소년업무규칙」 제31조부터 제33조에 따른 경찰청관서의 선도프로그램에 참여하는 경우',
    ],
  };
};

const createReportSection = (data: ReportData): (Paragraph | Table)[] => {
  let titleType = `(${data.type})`;
  if (data.type === '질병') titleType = '(질병)';
  else if (data.type === '출석인정') titleType = '(출석인정)';
  else if (data.type === '기타') titleType = '(기타)';

  const mainHeader = `결 석 신 고 서 (${data.type})`;
  const footerNotes = createFooterNoteParagraphs(data.type);

  return [
    // Top Header
    new Paragraph({
      children: [new TextRun({ text: `${data.schoolYear}학년도 ${titleType} 결석신고서 양식`, size: 16, font: 'Malgun Gothic' })],
      spacing: { after: 50 },
    }),

    // Approval Table (Top Right)
    new Table({
      alignment: AlignmentType.RIGHT,
      width: { size: 40, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: '결\n재', alignment: AlignmentType.CENTER, spacing: { line: 240 } })],
              verticalAlign: VerticalAlign.CENTER,
              width: { size: 15, type: WidthType.PERCENTAGE },
              rowSpan: 2,
              borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER },
            }),
            new TableCell({ 
              children: [new Paragraph({ text: '계', alignment: AlignmentType.CENTER })], 
              width: { size: 28, type: WidthType.PERCENTAGE },
              borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER },
            }),
            new TableCell({ 
              children: [new Paragraph({ text: '부장', alignment: AlignmentType.CENTER })], 
              width: { size: 28, type: WidthType.PERCENTAGE },
              borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER },
            }),
            new TableCell({ 
              children: [new Paragraph({ text: '교감', alignment: AlignmentType.CENTER })], 
              width: { size: 28, type: WidthType.PERCENTAGE },
              borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER },
            }),
          ],
        }),
        new TableRow({
          height: { value: 900, rule: HeightRule.ATLEAST },
          children: [
            new TableCell({ children: [], borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER } }),
            new TableCell({ children: [], borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER } }),
            new TableCell({ children: [], borders: { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER } }),
          ],
        }),
      ],
    }),

    new Paragraph({ spacing: { before: 200, after: 200 } }),

    // Main Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: mainHeader,
          bold: true,
          size: 48,
          font: 'Malgun Gothic',
        }),
      ],
    }),

    new Paragraph({ spacing: { before: data.type === '출석인정' ? 50 : 200, after: data.type === '출석인정' ? 50 : 200 } }),

    // Student Info Table (Aligned Right, No Borders for alignment)
    new Table({
      alignment: AlignmentType.RIGHT,
      width: { size: 45, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: '학 년 : ', size: 28, bold: true, font: 'Malgun Gothic' })], alignment: AlignmentType.LEFT })],
              width: { size: 30, type: WidthType.PERCENTAGE },
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: `제 ${data.year}학년 ${data.class}반 ${data.number}번`, size: 28, font: 'Malgun Gothic' })], alignment: AlignmentType.LEFT })],
              width: { size: 70, type: WidthType.PERCENTAGE },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: '성 명 : ', size: 28, bold: true, font: 'Malgun Gothic' })], alignment: AlignmentType.LEFT })],
              width: { size: 30, type: WidthType.PERCENTAGE },
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: data.name, size: 28, font: 'Malgun Gothic' })], alignment: AlignmentType.LEFT })],
              width: { size: 70, type: WidthType.PERCENTAGE },
            }),
          ],
        }),
      ],
    }),

    new Paragraph({ spacing: { before: 100, after: 200 } }),

    // Main Content
    new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({ text: `위 학생은 `, size: 26, font: 'Malgun Gothic' }),
        new TextRun({ text: `${data.reason}`, size: 26, bold: true, font: 'Malgun Gothic' }),
        new TextRun({ text: `(으)로 인하여 `, size: 26, font: 'Malgun Gothic' }),
        ...(data.startDate === data.endDate 
          ? [
              new TextRun({ text: `${data.startDate}`, size: 26, bold: true, font: 'Malgun Gothic' }),
            ]
          : [
              new TextRun({ text: `${data.startDate}`, size: 26, bold: true, font: 'Malgun Gothic' }),
              new TextRun({ text: ` ~ `, size: 26, font: 'Malgun Gothic' }),
              new TextRun({ text: `${data.endDate}`, size: 26, bold: true, font: 'Malgun Gothic' }),
            ]
        ),
        new TextRun({ text: ` (${data.days}일간) `, size: 26, bold: true, font: 'Malgun Gothic' }),
        new TextRun({ text: `결석하였기에 보호자 연서로 이 결석계를 제출합니다.`, size: 26, font: 'Malgun Gothic' }),
      ],
      spacing: { line: data.type === '출석인정' ? 360 : 400, after: data.type === '출석인정' ? 100 : 150 },
      indent: { firstLine: 400 },
    }),

    // Submission Date
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: data.submissionDate, size: 26, font: 'Malgun Gothic' })],
      spacing: { before: 100, after: 200 },
    }),

    // Signatures
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({ text: '학  생 : ', size: 26, font: 'Malgun Gothic' }),
        new TextRun({ text: data.name, size: 26, font: 'Malgun Gothic', bold: true }),
        new TextRun({ text: ' (인 또는 서명)', size: 18, color: '888888', font: 'Malgun Gothic' }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({ text: '보호자 : ', size: 26, font: 'Malgun Gothic' }),
        new TextRun({ text: (data.guardianName && data.guardianName.trim()) ? data.guardianName : '_______________', size: 26, font: 'Malgun Gothic', bold: true }),
        new TextRun({ text: ' (인 또는 서명)', size: 18, color: '888888', font: 'Malgun Gothic' }),
      ],
      spacing: { after: data.type === '출석인정' ? 100 : 400 },
    }),

    // Confirmation Box
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: '확 인 서', bold: true, size: 36, font: 'Malgun Gothic' })],
                  spacing: { before: data.type === '출석인정' ? 50 : 100, after: data.type === '출석인정' ? 100 : 200 },
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: '위 내용을 확인하였으며, 사실과 다름없습니다.', size: 26, font: 'Malgun Gothic' })],
                  spacing: { after: data.type === '출석인정' ? 100 : 200 },
                }),
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: data.submissionDate, size: 26, font: 'Malgun Gothic' })],
                  spacing: { after: 100 },
                }),
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun({ text: `담임교사 : ${data.teacherName} (인)`, size: 26, font: 'Malgun Gothic' }),
                  ],
                  spacing: { after: 100 },
                }),
              ],
              borders: {
                top: SINGLE_BORDER,
                bottom: SINGLE_BORDER,
                left: SINGLE_BORDER,
                right: SINGLE_BORDER,
              },
              margins: { top: 100, bottom: 100, left: 400, right: 400 },
            }),
          ],
        }),
      ],
    }),

    new Paragraph({ spacing: { before: data.type === '출석인정' ? 0 : 200, after: data.type === '출석인정' ? 0 : 200 } }),

    // School Name
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: '한올고등학교장 귀하', bold: true, size: 40, font: 'Malgun Gothic' })],
    }),

    new Paragraph({ spacing: { before: data.type === '출석인정' ? 0 : 200 } }),

    // Footer Notes
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: footerNotes.title, size: 14, font: 'Malgun Gothic', bold: true })],
                  spacing: { after: 30 },
                }),
                ...footerNotes.paragraphs.map((paragraph, index) => (
                  new Paragraph({
                    children: [new TextRun({ text: paragraph, size: 12, font: 'Malgun Gothic' })],
                    spacing: { after: index === footerNotes.paragraphs.length - 1 ? 0 : 20 },
                  })
                )),
              ],
              borders: {
                top: SINGLE_BORDER,
                bottom: SINGLE_BORDER,
                left: SINGLE_BORDER,
                right: SINGLE_BORDER,
              },
              margins: { top: 50, bottom: 50, left: 200, right: 200 },
            }),
          ],
        }),
      ],
    }),
  ];
};

export const generateAttendanceReport = async (data: ReportData) => {
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 400,
              bottom: 400,
              left: 1134,
              right: 1134,
            },
          },
        },
        children: createReportSection(data),
      },
    ],
  });

  return Packer.toBlob(doc);
};

export const generateBulkAttendanceReport = async (reports: ReportData[]) => {
  const doc = new Document({
    sections: reports.map(data => ({
      properties: {
        page: {
          margin: {
            top: 400,
            bottom: 400,
            left: 1134,
            right: 1134,
          },
        },
      },
      children: createReportSection(data),
    })),
  });

  return Packer.toBlob(doc);
};
