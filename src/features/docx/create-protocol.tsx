import { AlignmentType, Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, UnderlineType, VerticalAlign, WidthType } from "docx";
import { saveAs } from 'file-saver';
import { getDataProtocol } from "../data";

const data = getDataProtocol()
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
interface AnimalData {
    id: string;
    sampleId: string;
    ryr1: string;
    esr: string;
    igf2: string;
}
const animalData: AnimalData[] = [
    { id: "BY2000038492I3", sampleId: "c 10110", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849222", sampleId: "c 10111", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY20000384923l", sampleId: "c 10112", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849240", sampleId: "c 10113", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849l70", sampleId: "c 10114", ryr1: "NN", esr: "AA", igf2: "QQ" },
    { id: "BY200003849l 89", sampleId: "c 10115", ryr1: "NN", esr: "AA", igf2: "Qq" },
    { id: "BY200003849198", sampleId: "c 10116", ryr1: "NO", esr: "AA", igf2: "Qq" },
    { id: "BY200003849204", sampleId: "c 10117", ryr1: "NN", esr: "AB", igf2: "Qq" },
    { id: "BY200003849295", sampleId: "c 10118", ryr1: "NN", esr: "BB", igf2: "QQ" },
    { id: "BY20000384930l", sampleId: "c 10119", ryr1: "NN", esr: "AB", igf2: "QQ" },
    { id: "BY200003849310", sampleId: "c 10120", ryr1: "NN", esr: "AB", igf2: "QQ" },
    { id: "BY200003849329", sampleId: "c 10121", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849259", sampleId: "c 10122", ryr1: "NN", esr: "AA", igf2: "Qq" },
    { id: "BY200003849268", sampleId: "c 10123", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849277", sampleId: "c 10124", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY200003849286", sampleId: "c 10125", ryr1: "NN", esr: "AA", igf2: "Qq" },
    { id: "BY200003849453", sampleId: "c 10126", ryr1: "NN", esr: "AA", igf2: "Qq" },
    { id: "BY200003849462", sampleId: "c 10127", ryr1: "NN", esr: "AA", igf2: "qq" },
    { id: "BY20000384947 l", sampleId: "c 10128", ryr1: "NN", esr: "AA", igf2: "qq" },
];
const createProtocol = async () => {


    const protocol = new Document({
        sections: [
            {
                properties: {
                    page: {
                        margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
                    }
                },
                children: [
                    ...createHeader(),
                    ...createProtocolTitle(),
                    ...createProtocolInfo(),
                    new Paragraph({
                        pageBreakBefore: true
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({
                                text: 'Применяемые СИ, ВО и ИО в лаборатории*',
                                size: 24
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 400 },
                        children: [
                            new TextRun({
                                text: '*Указываются и предоставляются по требованию Заказчика',
                                size: 20
                            })
                        ]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({
                                text: 'Результаты исследований**: идентификационные № № образцов с 10110 – с 10128',
                                size: 24
                            }),
                        ],
                    }),
                    new Table({
                        rows: createTableProtocol(),
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
                            }),
                        ],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { before: 400 },
                        children: [
                            new TextRun({ text: 'Исполнители:', bold: true }),
                        ],
                    }),
                    new Table({
                        rows: createExecutorsSection(),
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: "none", size: 0, color: "FFFFFF" },
                            bottom: { style: "none", size: 0, color: "FFFFFF" },
                            left: { style: "none", size: 0, color: "FFFFFF" },
                            right: { style: "none", size: 0, color: "FFFFFF" },
                            insideHorizontal: { style: "none", size: 0, color: "FFFFFF" },
                            insideVertical: { style: "none", size: 0, color: "FFFFFF" },
                        },
                    }),
                    // ...createExecutorsSection(),
                    ...createFooter()
                ]

            },
        ],
    });

    const blob = await Packer.toBlob(protocol)
    saveAs(blob, 'protocol.docx')
}

// Функция для создания раздела с исполнителями
const createExecutorsSection = () => {
    return [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ text: 'Ведущий научный сотрудник', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                    width: { size: 40, type: WidthType.PERCENTAGE, }
                }),
                new TableCell({
                    children: [new Paragraph({ text: '10.02.2024', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                    width: { size: 20, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph({ text: 'Л.Н. Радюк', alignment: AlignmentType.RIGHT, spacing: { after: 400 } })],
                    width: { size: 20, type: WidthType.PERCENTAGE }
                })
            ]
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ text: 'Ведущий научный сотрудник', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                    width: { size: 40, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph({ text: '10.02.2024', alignment: AlignmentType.LEFT, spacing: { after: 400 } })],
                    width: { size: 20, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph({ text: 'Л.Н. Радюк', alignment: AlignmentType.RIGHT, spacing: { after: 400 } })],
                    width: { size: 20, type: WidthType.PERCENTAGE }
                })
            ],
        }),
    ];
};

// Функция для создания подвала документа
const createFooter = () => {
    return [
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 200 },
            children: [
                new TextRun({
                    text: `Дата выдачи протокола: ${formatDateToDMY(data.createdAt)}`,
                }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 400 },
            children: [
                new TextRun({
                    text: 'КОНЕЦ ПРОТОКОЛА',
                    bold: true,
                }),
            ],
        }),
        // new Paragraph({
        //     alignment: AlignmentType.RIGHT,
        //     children: [
        //         new TextRun({
        //             text: `п. 7.8 РК 420/11-03-02 Протокол № ${data.id}    Страница 1 из 2`,
        //             size: 20,
        //         }),
        //     ],
        // }),
    ];
};

const createTableProtocol = () => {
    return [
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
                    children: [new Paragraph({ text: "1", alignment: AlignmentType.CENTER, })],
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
        ...animalData.map((item, index) => (
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
    ]
}



// Функция для создания основной информации протокола
const createProtocolInfo = () => {

    const SIZE_INFO = 24
    return [
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Дата доставки проб на исследования: `,
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `dd.mm.year`,
                    shading: {
                        color: "00FFFF",
                        fill: "FF0000",
                    },
                    size: SIZE_INFO
                }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: 'Наименование объекта исследований: ',
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `образцы ткани свиней`,
                    size: SIZE_INFO,
                    shading: {
                        color: "00FFFF",
                        fill: "FF0000",
                    }
                })
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Количество исследуемых образцов: `,
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `кол-во`,
                    size: SIZE_INFO,
                    shading: {
                        color: "00FFFF",
                        fill: "FF0000",
                    },
                }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Наименование организации, производившей отбор образцов: ${data.customer.orgInformations.fullName}, ${data.customer.orgInformations.address}`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Наименование организации, производившей отбор образцов: ${data.customer.orgInformations.fullName}, ${data.customer.orgInformations.address}`,
                    size: SIZE_INFO,
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `заявка (сопроводительная) от ${formatDateToDMY(data.createdAt)}`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Наименование организации-заказчика: ${data.customer.orgInformations.fullName}, ${data.customer.orgInformations.address}`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: 'Вид исследований: ',
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `ДНК-тестирование по генам, контролирующим хозяйственнозначимые признаки и детерминирующим наследственные заболевания`,
                    size: SIZE_INFO,
                    shading: {
                        color: "00FFFF",
                        fill: "FF0000",
                    },
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: 'Наименование методики, устанавливающей метод исследований: ',
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `«Методические рекомендации по применению метода ДНК-диагностики наследственных
заболеваний и генетической устойчивости свиней к инфекционным заболеваниям»
(утверждены Генеральным директором РУП «НПЦ НАН Беларуси по животноводству»,
протокол № 22 от 22 ноября 2013 г., рассмотрены и одобрены на заседании секции
животноводства и ветеринарии научно-технического совета Министерства сельского
хозяйства и продовольствия Республики Беларусь, протокол № 13 от 17 июня 2014 г.)`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Дата начала исследований: ${formatDateToDMY(data.createdAt)}`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Дата начала исследований: ${formatDateToDMY(data.createdAt)}`,
                    size: SIZE_INFO
                })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 200, after: 200 },
            children: [
                new TextRun({
                    text: `Условия проведения исследований: температура воздуха: `,
                    size: SIZE_INFO
                }),
                new TextRun({
                    text: `24,0-27,9°С, относительная влажность 20,7-36,5%`,
                    size: SIZE_INFO
                })
            ]
        }),
    ];
};



// Функция для создания заголовка протокола
const createProtocolTitle = () => {
    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 1000 },
            children: [
                new TextRun({
                    text: `ПРОТОКОЛ ИССЛЕДОВАНИЙ № ${data.id}`,
                    bold: true,
                    size: 28,
                }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 1000 },
            children: [
                new TextRun({
                    text: `от ${formatDateToDMY(data.createdAt)}`,
                    size: 24
                }),
            ],
        }),
    ];
};

const createHeader = () => {

    const SIZE_FONT_LEFT_CELL = 20

    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: {
                after: 300
            },
            children: [
                new TextRun({
                    text: data.executor.orgInformations.fullName.toUpperCase(),
                    bold: true,
                    size: 24,

                }),
            ],
        }),
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE,
            },
            columnWidths: [5500, 2000, 2500], // Равные колонки
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
                        // Левая колонка
                        new TableCell({
                            width: {
                                size: 5500,
                                type: WidthType.DXA,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `${data.executor.accredits[0].accredited}`,
                                            size: SIZE_FONT_LEFT_CELL
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `Аттестат аккредитации ${data.executor.accredits[0].certificate}`,
                                            size: SIZE_FONT_LEFT_CELL
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: data.executor.orgInformations.address,
                                            size: SIZE_FONT_LEFT_CELL
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `т./факс ${data.executor.orgContacts[0].telephone}`,
                                            size: SIZE_FONT_LEFT_CELL
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,

                                    children: [
                                        new TextRun({
                                            text: 'укажите ваш веб-сайт',
                                            size: SIZE_FONT_LEFT_CELL,
                                            shading: {
                                                color: "00FFFF",
                                                fill: "FF0000",
                                            },
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [
                                        new TextRun({
                                            text: `E-mail: ${data.executor.orgContacts[0].web}`,
                                            size: SIZE_FONT_LEFT_CELL
                                        }),
                                    ],
                                }),
                            ],
                        }),
                        new TableCell({
                            width: {
                                size: 2000,
                                type: WidthType.DXA,
                            },
                            children: [
                                // Можно добавить логотип или оставить пустым
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                        new TextRun({
                                            text: "", // Место для логотипа
                                        }),
                                    ],
                                }),
                            ],
                        }),

                        // Правая колонка (утверждающая часть)
                        new TableCell({
                            width: {
                                size: 2500,
                                type: WidthType.DXA,
                            },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    spacing: { before: 0, after: 100 },
                                    children: [
                                        new TextRun({
                                            text: 'УТВЕРЖДАЮ',
                                            bold: true,
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.JUSTIFIED,
                                    spacing: { before: 0, after: 100 },
                                    children: [
                                        new TextRun({
                                            text: 'Заведующий лабораторией',
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.JUSTIFIED,
                                    spacing: { before: 0, after: 100 },
                                    children: [
                                        new TextRun({
                                            text: '_________ В.П. Симоненко',
                                        }),
                                    ],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.JUSTIFIED,
                                    spacing: { before: 0, after: 0 },
                                    children: [
                                        new TextRun({
                                            text: '«   » ',
                                            underline: {
                                                type: UnderlineType.SINGLE,
                                            },
                                        }),
                                        new TextRun({
                                            text: '  '
                                        }),
                                        new TextRun({
                                            text: '_____________'
                                        }),
                                        new TextRun({
                                            text: ' '
                                        }),
                                        new TextRun({
                                            text: '2024 г.'
                                        })
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        }),
    ]
}

export function ButtonCreateProtocol() {
    return <button onClick={createProtocol}>протокол</button>
}
