import React from 'react';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, TextRun, AlignmentType, VerticalAlign } from 'docx';

interface AnimalData {
    id: string;
    sampleId: string;
    ryr1: string;
    esr: string;
    igf2: string;
}

interface TableGeneratorProps {
    data: AnimalData[];
    fileName?: string;
}

const TableGenerator: React.FC<TableGeneratorProps> = ({ data, fileName = 'protocol' }) => {
    const generateDocument = async () => {
        // Create table rows from data
        const tableRows = [
            // Main header row (merged horizontally)
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ text: "№ п.п.", alignment: AlignmentType.CENTER })],
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        rowSpan: 3, // Объединяем вертикально 3 строки
                        verticalAlign: VerticalAlign.CENTER
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "Индивидуальный № животного", alignment: AlignmentType.CENTER })],
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        rowSpan: 3, // Объединяем вертикально 3 строки
                        verticalAlign: VerticalAlign.CENTER
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "Идентификационный № образца", alignment: AlignmentType.CENTER })],
                        width: { size: 15, type: WidthType.PERCENTAGE },
                        rowSpan: 3, // Объединяем вертикально 3 строки
                        verticalAlign: VerticalAlign.CENTER
                    }),
                    new TableCell({
                        children: [new Paragraph({
                            text: "Исследуемые показатели (хозяйственно-ценные признаки и наследственные заболевания)",
                            alignment: AlignmentType.CENTER
                        })],
                        columnSpan: 3,
                    }),
                ],
            }),
            // Subheader row 1 - "Исследуемый ген и генотип животного"
            new TableRow({
                children: [
                    // Первые 3 ячейки объединены rowSpan в предыдущей строке
                    new TableCell({
                        children: [new Paragraph({
                            text: "Исследуемый ген и генотип животного",
                            alignment: AlignmentType.CENTER
                        })],
                        columnSpan: 3,
                    }),
                ],
            }),
            // Subheader row 2 - Gene names (RYR1, ESR, IGF2)
            new TableRow({
                children: [
                    // Первые 3 ячейки объединены rowSpan
                    new TableCell({
                        children: [new Paragraph({ text: "RYR1", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "ESR", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "IGF2", alignment: AlignmentType.CENTER })],
                    }),
                ],
            }),
            // Subheader row 3 - Column numbers (*4*, *5*, *6*)
            new TableRow({
                children: [
                    // Первые 3 ячейки объединены rowSpan
                    new TableCell({
                        children: [new Paragraph({ text: "1", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "2", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "3", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "4", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "5", alignment: AlignmentType.CENTER })],
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "6", alignment: AlignmentType.CENTER })],
                    }),
                ],
            }),
            // Data rows
            ...data.map((item, index) => (
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ text: (index + 1).toString(), alignment: AlignmentType.CENTER })],
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: item.id, alignment: AlignmentType.CENTER })],
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: item.sampleId, alignment: AlignmentType.CENTER })],
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: item.ryr1, alignment: AlignmentType.CENTER })],
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: item.esr, alignment: AlignmentType.CENTER })],
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: item.igf2, alignment: AlignmentType.CENTER })],
                        }),
                    ],
                })
            )),
        ];

        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Протокол генетических исследований",
                                bold: true,
                                size: 28,
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 400 },
                    }),
                    new Table({
                        rows: tableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: "single", size: 4, color: "000000" },
                            bottom: { style: "single", size: 4, color: "000000" },
                            left: { style: "single", size: 4, color: "000000" },
                            right: { style: "single", size: 4, color: "000000" },
                            insideHorizontal: { style: "single", size: 2, color: "000000" },
                            insideVertical: { style: "single", size: 2, color: "000000" },
                        },
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 200 },
                        children: [
                            new TextRun({
                                text: '**Результаты исследований распространяются только на исследованные образцы, предоставленные заказчиком, который несет ответственность за правильность отбора образцов',
                                italics: true,
                            }),
                        ],
                    }),
                ],
            }],
        });

        // Generate the DOCX file and trigger download
        const blob = await Packer.toBlob(doc);
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${fileName}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    return (
        <button
            onClick={generateDocument}
            style={{
                padding: '10px 20px',
                backgroundColor: '#4CAF50',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '16px',
            }}
        >
            Сгенерировать протокол
        </button>
    );
};

export default TableGenerator;