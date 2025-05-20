import React from 'react';
import { 
  Document, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  AlignmentType, 
  WidthType,
  VerticalAlign,
  Packer,
  BorderStyle,
  HeadingLevel,
  UnderlineType
} from 'docx';
import { saveAs } from 'file-saver';

interface ApplicationData {
  applicationNumber: number;
  applicationType: 'diseases' | 'origin'; // Тип заявки: на заболевания или происхождение
  applicant: string;
  legalAddress: string;
  phone: string;
  director: string;
  basis: string;
  bankDetails: string;
  samplesType: string;
  contactPerson: string;
  animals: AnimalData[];
  selectionDate?: string;
  selectionPlace?: string;
  selectionDocument?: string;
  deliveryDate?: string;
  deliveryMethod?: string;
  selectionEmployee?: string;
  receivingEmployee?: string;
}

interface AnimalData {
  id: number;
  animalNumber: string;
  birthDate?: string;
  motherInfo?: string;
  fatherInfo?: string;
  category?: string;
  breed?: string;
  tests: {
    CVM: boolean;
    BLAD: boolean;
    BY: boolean;
    DUMPS: boolean;
    BC: boolean;
    FXID: boolean;
    HH1: boolean;
    HH3: boolean;
    HH4: boolean;
    HH5: boolean;
    HCD: boolean;
    geneticExpertise: boolean;
  };
  sampleNumber?: string;
  sampleType?: string;
  testType?: 'diseases' | 'origin';
}

const ApplicationGenerator: React.FC<{ data: ApplicationData }> = ({ data }) => {
  const generateDocument = () => {
    // Создаем таблицу для заявки на заболевания
    const createDiseasesTable = () => {
      const headerRow = new TableRow({
        children: [
          createHeaderCell('№ п/п', 5),
          createHeaderCell('Индивидуальный номер животного', 13),
          createHeaderCell('Дата рождения бычка', 8),
          createHeaderCell('Инд. номер и кличка матери', 17),
          createHeaderCell('Инд. номер и кличка отца', 9),
          ...['CVM', 'BLAD', 'BY', 'DUMPS', 'BC', 'FXID', 'HH1', 'HH3', 'HH4', 'HH5', 'HCD'].map(test => 
            createHeaderCell(test, 4)
          )
        ]
      });

      const subHeaderRow = new TableRow({
        children: [
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('Необходимые исследования', 44, false, 11)
        ]
      });

      const dataRows = data.animals.map(animal => (
        new TableRow({
          children: [
            createDataCell(animal.id.toString(), 5),
            createDataCell(animal.animalNumber, 13),
            createDataCell(animal.birthDate || '', 8),
            createDataCell(animal.motherInfo || '', 17),
            createDataCell(animal.fatherInfo || '', 9),
            createDataCell(animal.tests.CVM ? '+' : '', 4),
            createDataCell(animal.tests.BLAD ? '+' : '', 4),
            createDataCell(animal.tests.BY ? '+' : '', 4),
            createDataCell(animal.tests.DUMPS ? '+' : '', 4),
            createDataCell(animal.tests.BC ? '+' : '', 4),
            createDataCell(animal.tests.FXID ? '+' : '', 4),
            createDataCell(animal.tests.HH1 ? '+' : '', 4),
            createDataCell(animal.tests.HH3 ? '+' : '', 4),
            createDataCell(animal.tests.HH4 ? '+' : '', 4),
            createDataCell(animal.tests.HH5 ? '+' : '', 4),
            createDataCell(animal.tests.HCD ? '+' : '', 4)
          ]
        })
      ));

      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [5, 13, 8, 17, 9, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4],
        rows: [headerRow, subHeaderRow, ...dataRows],
        borders: createTableBorders()
      });
    };

    // Создаем таблицу для заявки на происхождение
    const createOriginTable = () => {
      const headerRow = new TableRow({
        children: [
          createHeaderCell('№ п/п', 10),
          createHeaderCell('Инд. номер коровы, кличка, порода', 20),
          createHeaderCell('Категория животного', 15),
          createHeaderCell('Дата рождения', 15),
          createHeaderCell('Инд. номер и кличка матери', 20),
          createHeaderCell('Инд. номер и кличка отца', 20),
          createHeaderCell('Необходимые исследования', 20)
        ]
      });

      const subHeaderRow = new TableRow({
        children: [
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('', 0, true),
          createHeaderCell('Генетическая экспертиза', 20)
        ]
      });

      const dataRows = data.animals.map(animal => (
        new TableRow({
          children: [
            createDataCell(animal.id.toString(), 10),
            createDataCell(`${animal.animalNumber} ${animal.breed || ''}`, 20),
            createDataCell(animal.category || '', 15),
            createDataCell(animal.birthDate || '', 15),
            createDataCell(animal.motherInfo || '', 20),
            createDataCell(animal.fatherInfo || '', 20),
            createDataCell(animal.tests.geneticExpertise ? '+' : '', 20)
          ]
        })
      ));

      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [10, 20, 15, 15, 20, 20, 20],
        rows: [headerRow, subHeaderRow, ...dataRows],
        borders: createTableBorders()
      });
    };

    // Создаем таблицу для акта отбора
    const createSelectionActTable = () => {
      const headerRow = new TableRow({
        children: [
          createHeaderCell('№ п/п', 5),
          createHeaderCell('№ пробы', 10),
          createHeaderCell('Индивидуальный номер, кличка животного', 20),
          createHeaderCell('Порода', 15),
          createHeaderCell('Дата рождения', 10),
          createHeaderCell('Вид биологического материала', 20),
          createHeaderCell('Вид генетического тестирования', 20)
        ]
      });

      const dataRows = data.animals.map(animal => (
        new TableRow({
          children: [
            createDataCell(animal.id.toString(), 5),
            createDataCell(animal.sampleNumber || '', 10),
            createDataCell(animal.animalNumber, 20),
            createDataCell(animal.breed || '', 15),
            createDataCell(animal.birthDate || '', 10),
            createDataCell(animal.sampleType || '', 20),
            createDataCell(
              animal.testType === 'diseases' 
                ? 'Выявление генетически-детерминированных заболеваний' 
                : 'Генетическая экспертиза на достоверность происхождения', 
              20
            )
          ]
        })
      ));

      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [5, 10, 20, 15, 10, 20, 20],
        rows: [headerRow, ...dataRows],
        borders: createTableBorders()
      });
    };

    // Создаем документ
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1000,
              right: 1000,
              bottom: 1000,
              left: 1000,
            },
          },
        },
        children: [
          // Заявка №1
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ 
              text: `ЗАЯВКА №${data.applicationNumber}`, 
              bold: true,
              allCaps: true
            })],
          }),
          
          new Paragraph({
            children: [
              new TextRun({
                text: 'по генетической экспертизе поголовья на ',
              }),
              new TextRun({
                text: 'выявление детерминированных (наследственных) заболеваний',
                bold: true
              }),
              new TextRun({
                text: ' ремонтных бычков молочных пород от быкопроизводящих коров племенных хозяйств, племенных быков селекционно-генетических центров'
              })
            ]
          }),

          new Paragraph({
            text: 'Заявитель на проведение исследования:',
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: data.applicant,
            indent: { left: 2000 }
          }),

          new Paragraph({
            text: `Вид образцов: ${data.samplesType}${'\t'.repeat(10)}(нужное подчеркнуть)`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: 'Юридический адрес:',
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: data.legalAddress,
            indent: { left: 2000 }
          }),

          new Paragraph({
            text: `Телефон (факс): ${data.phone}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `ФИО и должность руководителя (полностью): ${data.director}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `действующего на основании: ${data.basis}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Банковские реквизиты, адрес банка: ${data.bankDetails}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: 'Сопроводительный документ',
            bold: true,
            spacing: { before: 400 }
          }),

          data.applicationType === 'diseases' ? createDiseasesTable() : createOriginTable(),

          new Paragraph({
            text: `Дата: ${' '.repeat(30)}${data.selectionDate || '_'.repeat(10)}`,
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: `Контактное лицо: ${data.contactPerson}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: 'Подпись, печать',
            italics: true,
            alignment: AlignmentType.RIGHT,
            spacing: { before: 400 }
          }),

          // Акт №1
          new Paragraph({
            text: 'АКТ №1',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 800 },
            pageBreakBefore: true
          }),

          new Paragraph({
            text: 'Отбора биологического материала',
            alignment: AlignmentType.CENTER,
            bold: true
          }),

          new Paragraph({
            text: `Название организации: ${data.applicant}`,
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: `Вид животных: ${data.animals[0]?.breed || ''}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Место проведения отбора: ${data.selectionPlace || ''}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Дата отбора образцов: ${data.selectionDate || ''}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Наименование документа, в соответствии с которым произведён отбор: ${data.selectionDocument || ''}`,
            spacing: { before: 200 }
          }),

          createSelectionActTable(),

          new Paragraph({
            text: `Отбор произведен в емкость ${data.samplesType}`,
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: `Материал отправлен (дата): ${data.deliveryDate || ''}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Способ доставки: ${data.deliveryMethod || ''}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: `Сотрудник, производивший отбор: ${data.selectionEmployee || '_'.repeat(30)}`,
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: '(ФИО, должность, тел.)',
            italics: true,
            indent: { left: 2000 }
          }),

          new Paragraph({
            text: `Контактное лицо: ${data.contactPerson || '_'.repeat(30)}`,
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: '(ФИО, тел.)',
            italics: true,
            indent: { left: 2000 }
          }),

          new Paragraph({
            text: 'В лаборатории генетики животных Института генетики и цитологии НАН Беларуси образцы принял:',
            spacing: { before: 400 }
          }),

          new Paragraph({
            text: data.receivingEmployee || '_'.repeat(80),
            spacing: { before: 200 }
          }),

          new Paragraph({
            text: 'ФИО, должность сотрудника, производившего приемку образцов, дата',
            italics: true,
            indent: { left: 2000 }
          })
        ]
      }]
    });

    Packer.toBlob(doc).then(blob => {
      saveAs(blob, `Заявка_акт_ФОРМА_${data.applicationNumber}.docx`);
    });
  };

  // Вспомогательные функции
  const createHeaderCell = (text: string, widthPercent: number, skipRender = false, colSpan = 1) => {
    if (skipRender) return new TableCell({});
    
    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      columnSpan: colSpan,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: true })],
        }),
      ],
      verticalAlign: VerticalAlign.CENTER,
      shading: { fill: 'DDDDDD', type: 'clear' }
    });
  };

  const createDataCell = (text: string, widthPercent: number) => {
    return new TableCell({
      width: { size: widthPercent, type: WidthType.PERCENTAGE },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text })],
        }),
      ],
      verticalAlign: VerticalAlign.CENTER
    });
  };

  const createTableBorders = () => ({
    top: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    left: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    right: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
  });

  return (
    <div style={{ padding: '20px' }}>
      <button 
        onClick={generateDocument}
        style={{
          padding: '10px 20px',
          fontSize: '16px',
          backgroundColor: '#4CAF50',
          color: 'white',
          border: 'none',
          borderRadius: '4px',
          cursor: 'pointer'
        }}
      >
        Сгенерировать документ
      </button>
    </div>
  );
};

export default ApplicationGenerator;