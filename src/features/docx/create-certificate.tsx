import { AlignmentType, BorderStyle, Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType } from "docx";
import { saveAs } from 'file-saver';
import { getDataCertificate } from "../data";

const data = getDataCertificate()
interface ResearchMarker {
    id: number;
    researchMarker: {
        marker: {
            name: string;
        };
        description: string;
    };
}
function formatDateToDMY(isoDateString: string) {
    // Создаем объект Date из ISO строки
    const date = new Date(isoDateString);

    // Получаем компоненты даты
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Месяцы 0-11
    const year = date.getFullYear();

    // Форматируем в dd.mm.yyyy
    return `${day}.${month}.${year}`;
}

const createCertificate = async () => {

    console.log(getDataCertificate());

    const { table1, table2 } = splitMarkers();


    const protocol = new Document({
        sections: [
            {
                properties: {},
                children: [
                    ...createHeader(),
                    ...createCertificateTitle(),
                    ...createCertificateAnimalInfo(),
                    new Paragraph({
                        spacing: { after: 200 },
                    }),
                    createMarkerTable(table1),
                    new Paragraph({
                        spacing: { before: 100 },
                    }),
                    createMarkerTable(table2),
                    new Paragraph({
                        spacing: { before: 100 },
                    }),
                    ...createConfirmOrigin(),
                    ...createCertificateInfo(),
                    ...createExecutorSection(),
                    new Paragraph({
                        spacing: { before: 200 },
                    }),
                    ...createFooterCertificate()
                ],
            },
        ],
    });

    const blob = await Packer.toBlob(protocol)
    saveAs(blob, `Генетический сертификат №${data.id}.docx`)
}
//Функция создания выдачи сертификата
const createFooterCertificate = () => {
    return [
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE
            },
            borders: {
                top: { style: "none", size: 0, color: "FFFFFF" },
                bottom: { style: "none", size: 0, color: "FFFFFF" },
                left: { style: "none", size: 0, color: "FFFFFF" },
                right: { style: "none", size: 0, color: "FFFFFF" },
                insideHorizontal: { style: "none", size: 0, color: "FFFFFF" },
                insideVertical: { style: "none", size: 0, color: "FFFFFF" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ text: 'Дата выдачи сертификата', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                            width: { size: 40, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: '10.02.2024', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                            width: { size: 20, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: 'М.П.', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                            width: { size: 20, type: WidthType.PERCENTAGE }
                        })
                    ],
                }),
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 600 },
            children: [
                new TextRun({
                    text: `Документ оформлен в 2-х экземплярах, не допускается копирование и тиражирование полностью или частично`,
                    size: 12
                })
            ]
        })
    ]
}

// Функция для создания раздела с исполнителями

const createExecutorSection = () => {

    const WIDTH_FIRST_CELL = 30
    const WIDTH_SECOND_CELL = 45
    const WIDTH_THIRD_CELL = 15

    return [
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE
            },
            borders: {
                top: { style: "none", size: 0, color: "FFFFFF" },
                bottom: { style: "none", size: 0, color: "FFFFFF" },
                left: { style: "none", size: 0, color: "FFFFFF" },
                right: { style: "none", size: 0, color: "FFFFFF" },
                insideHorizontal: { style: "none", size: 0, color: "FFFFFF" },
                insideVertical: { style: "none", size: 0, color: "FFFFFF" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: WIDTH_FIRST_CELL,
                                type: WidthType.PERCENTAGE,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: `Ответственный за проведение генетической экспертизы / сертификат оформил младший научный сотрудник`,
                                            size: 20
                                        }),
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_SECOND_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({ text: ' ', spacing: { after: 200 } }),
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_THIRD_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            verticalAlign: VerticalAlign.BOTTOM,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: 'Л.В. Глущенко',
                                            size: 20
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: WIDTH_FIRST_CELL,
                                type: WidthType.PERCENTAGE,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: `Сертификат проверил техник 1 категории`,
                                            size: 20
                                        }),
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_SECOND_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({ text: ' ', spacing: { after: 200 } }),
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_THIRD_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            verticalAlign: VerticalAlign.BOTTOM,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: 'Л.Н. Радюк',
                                            size: 20
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: WIDTH_FIRST_CELL,
                                type: WidthType.PERCENTAGE,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: `Заведующий лабораторией`,
                                            size: 20
                                        }),
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_SECOND_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({ text: ' ', spacing: { after: 200 } }),
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: WIDTH_THIRD_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            verticalAlign: VerticalAlign.BOTTOM,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 200 },
                                    children: [
                                        new TextRun({
                                            text: 'В.П. Симоненко',
                                            size: 20
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
            ]
        })
    ]
}

// Функция создания информации сертификата
const createCertificateInfo = () => {
    return [
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 400 },
            children: [
                new TextRun({
                    text: 'Полученные результаты относятся к образцу, предоставленному заказчиком:',
                    size: 20
                }),
                new TextRun({
                    text: ` ${data.customer.orgInformations.fullName}, ${data.customer.orgInformations.address}`
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Генетический сертификат составлен на основе протокола № 1462`,
                    size: 18,
                    italics: true
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
                new TextRun({
                    text: 'Вид исследований / ТНПА на метод исследований:',
                    size: 18
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { after: 400 },
            children: [
                new TextRun({
                    text: `- оценка достоверности происхождения сельскохозяйственных животных по полиморфизму нуклеотидных
последовательностей ДНК / «Методические рекомендации по проведению генотипирования крупного рогатого скота
по микросателлитным локусам ДНК» (одобрены и рекомендованы к использованию НТС Минсельхозпрода Республики
Беларусь, протокол № 22 от 20 февраля 2015 г.)`,
                    size: 18,
                    italics: true
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Аллели приведены в соответствие к стандарту ISAG (тест сравнения 2016/17г.)`,
                    italics: true
                })
            ]
        })
    ]
}

// Фунцкция создания подвтверждения происхождения
const createConfirmOrigin = () => {
    return [
        new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
                new TextRun({
                    text: 'Результаты исследований: \t',
                    size: 20
                }),
                new TextRun({
                    text: 'Отцовство — ',
                    size: 20
                }),
                new TextRun({
                    text: `подтвержается`,
                    size: 20
                }),
                new TextRun({
                    text: '\t Материнство — ',
                    size: 20
                }),
                new TextRun({
                    text: `подтвержается`,
                    size: 20
                }),
            ]
        })
    ]
}

//Функции создания таблицы маркеров
const splitMarkers = () => {
    const researchOnMarkers = data.researchs[0].researchOnMarkers

    const half = Math.ceil(researchOnMarkers.length / 2);
    return {
        table1: researchOnMarkers.slice(0, half),
        table2: researchOnMarkers.slice(half)
    };
};

const createMarkerTable = (markers: ResearchMarker[]) => {
    // Собираем массивы имён и значений
    const markerNames = markers.map(m => m.researchMarker.marker.name);
    const markerValues = markers.map(m => m.researchMarker.description);

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
            left: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
            right: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
            insideVertical: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
        },
        rows: [
            // Строка с названиями маркеров (жирный текст, серый фон)
            new TableRow({
                children: markerNames.map(name => (
                    new TableCell({
                        width: { size: 12.5, type: WidthType.PERCENTAGE },
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({
                                    text: name,
                                    bold: true
                                })],
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                        // shading: { fill: 'DDDDDD', type: ShadingType.CLEAR }
                    })
                ))
            }),
            // Строка со значениями маркеров
            new TableRow({
                children: markerValues.map(value => (
                    new TableCell({
                        width: { size: 12.5, type: WidthType.PERCENTAGE },
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: value })],
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER
                    })
                ))
            })
        ]
    });
};

// Функция создания инофрамции о животном в сертификате
const createCertificateAnimalInfo = () => {

    const SIZE_FIRST_TABLE_CELL = 50
    const SIZE_SECOND_TABLE_CELL = 20
    const SIZE_THIRD_TABLE_CELL = 26
    const FONT_SIZE = 20

    return [
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE
            },
            borders: {
                top: { style: "none", size: 0, color: "FFFFFF" },
                bottom: { style: "none", size: 0, color: "FFFFFF" },
                left: { style: "none", size: 0, color: "FFFFFF" },
                right: { style: "none", size: 0, color: "FFFFFF" },
                insideHorizontal: { style: "none", size: 0, color: "FFFFFF" },
                insideVertical: { style: "none", size: 0, color: "FFFFFF" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: SIZE_FIRST_TABLE_CELL,
                                type: WidthType.PERCENTAGE,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Идентификационный номер образца: `,
                                            size: FONT_SIZE
                                        }),
                                        // new TextRun({
                                        //     text: ` `,
                                        //     size: FONT_SIZE
                                        // }),
                                        new TextRun({
                                            text: `номер образца`,
                                            size: FONT_SIZE,
                                            shading: {
                                                color: "00FFFF",
                                                fill: "FF0000",
                                            },
                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: SIZE_SECOND_TABLE_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: ` `,

                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: SIZE_THIRD_TABLE_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun({
                                            text: 'Происхождение животного:',
                                            size: FONT_SIZE
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE
                            },
                            columnSpan: 3,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Идентификационный номер животного/кличка: ${data.researchs[0].animalResearch.indNumber} ${data.researchs[0].animalResearch.name}`,
                                            size: FONT_SIZE
                                        })
                                    ]
                                })
                            ]
                        }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            columnSpan: 2,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Пол: ${data.researchs[0].animalResearch.gender.name} \t\t\t Дата рождения: ${formatDateToDMY(data.researchs[0].animalResearch.dob)}`,
                                            size: FONT_SIZE
                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: SIZE_THIRD_TABLE_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun({
                                            text: `Отец: `,
                                            size: FONT_SIZE
                                        }),
                                        new TextRun({
                                            text: "—",
                                            size: FONT_SIZE,
                                            bold: true
                                        }),
                                        new TextRun({
                                            text: ` ${data.researchs[0].animalResearch.father.indNumber}`,
                                            size: FONT_SIZE,
                                        })
                                    ]
                                })
                            ]
                        }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE
                            },
                            columnSpan: 3,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Вид животного: крупный рогатый скот`,
                                            size: FONT_SIZE
                                        })
                                    ]
                                })
                            ]
                        }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            columnSpan: 2,
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Порода: ${data.researchs[0].animalResearch.breed.breed} `,
                                            size: FONT_SIZE
                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: {
                                size: SIZE_THIRD_TABLE_CELL,
                                type: WidthType.PERCENTAGE
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun({
                                            text: `Мать: `,
                                            size: FONT_SIZE
                                        }),
                                        new TextRun({
                                            text: "—",
                                            size: FONT_SIZE,
                                            bold: true
                                        }),
                                        new TextRun({
                                            text: ` ${data.researchs[0].animalResearch.mother.indNumber}`,
                                            size: FONT_SIZE,
                                        })
                                    ]
                                })
                            ]
                        }),
                    ]
                }),
            ]
        })
    ]
}

// Функция для создания заголовка протокола
const createCertificateTitle = () => {
    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400 },
            children: [
                new TextRun({
                    text: `ГЕНЕТИЧЕСКИЙ СЕРТИФИКАТ № ${data.id}`,
                    bold: true,
                    size: 28,
                }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
                new TextRun({
                    text: `ЭКЗЕМПЛЯР 1`,
                    size: 24
                }),
            ],
        }),
    ];
};

//Функция создания шапки документа сертификата
const createHeader = () => {
    return [
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE,
            },
            // columnWidths: [5500, 2000, 2500], 
            borders: {
                top: { style: "none", size: 0, color: "FFFFFF" },
                bottom: { style: "none", size: 0, color: "FFFFFF" },
                left: { style: "none", size: 0, color: "FFFFFF" },
                right: { style: "none", size: 0, color: "FFFFFF" },
                insideHorizontal: { style: "none", size: 0, color: "FFFFFF" },
                insideVertical: { style: "none", size: 0, color: "FFFFFF" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { after: 400 },

                                    children: [
                                        new TextRun({
                                            text: 'картинка',
                                            size: 20
                                        })
                                    ]
                                })
                            ],
                            verticalAlign: VerticalAlign.CENTER,
                            width: {
                                size: 20,
                                type: WidthType.PERCENTAGE,
                            }
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                        new TextRun({
                                            text: 'РУП «НАУЧНО-ПРАКТИЧЕСКИЙ ЦЕНТР НАН БЕЛАРУСИ ПО ЖИВОТНОВОДСТВУ»',
                                            size: 20,
                                            bold: true
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.BOTH,
                                    children: [
                                        new TextRun({
                                            text: `${data.executor.accredits[0].accredited}.`,
                                            size: 16,
                                        }),
                                        new TextRun({
                                            text: ` `,
                                            size: 16,
                                        }),
                                        new TextRun({
                                            text: `Аттестат аккредитации ${data.executor.accredits[0].certificate}. `,
                                            size: 16
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.BOTH,
                                    children: [
                                        new TextRun({
                                            text: `${data.executor.orgInformations.address}.`,
                                            size: 16,
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.BOTH,
                                    children: [
                                        new TextRun({
                                            text: `т./факс ${data.executor.orgContacts[0].telephone}.`,
                                            size: 16,
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.BOTH,
                                    children: [
                                        new TextRun({
                                            text: `Ваш веб-сайт`,
                                            size: 16,
                                            shading: {
                                                color: "00FFFF",
                                                fill: "FF0000",
                                            },
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.BOTH,
                                    children: [
                                        new TextRun({
                                            text: `E-mail: ${data.executor.orgContacts[0].web}`
                                        })
                                    ]
                                })
                            ],
                            width: {
                                size: 60,
                                type: WidthType.PERCENTAGE,
                            }
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    spacing: { after: 400 },

                                    children: [
                                        new TextRun({
                                            text: 'картинка',
                                            size: 20
                                        })
                                    ]
                                })
                            ],
                            verticalAlign: VerticalAlign.CENTER,
                            width: {
                                size: 30,
                                type: WidthType.PERCENTAGE,
                            }
                        }),
                    ]
                }),
            ]
        })
    ]
}

export function ButtonCreateCertificate() {
    return <button onClick={createCertificate}>сертификат</button>
}
