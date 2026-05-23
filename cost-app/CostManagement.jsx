import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const PROD_COST_PCT = 0.35;

const RAW_MATERIALS = [{"code": "PPR-001", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0631\u0627\u0646\u062f\u0648\u0645 R40", "lastCost": 100.0, "avgCost": 80.1845484377229, "itemType": "R"}, {"code": "PPH-001", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 H26ML", "lastCost": 59.0, "avgCost": 67.4507739938081, "itemType": "R"}, {"code": "PPK-001", "desc": "\u062e\u0627\u0645\u0629 \u0643\u0648\u0628\u0644\u0645\u0631", "lastCost": 115.0, "avgCost": 75.3691770186335, "itemType": "R"}, {"code": "ABS-001", "desc": "\u062e\u0627\u0645\u0629 ABS", "lastCost": 98.0, "avgCost": 106.131313131313, "itemType": "R"}, {"code": "PPS-001", "desc": "\u0628\u0648\u0644\u064a \u0627\u0633\u062a\u0631\u064a\u0646", "lastCost": 72.5, "avgCost": 86.5839025839026, "itemType": "R"}, {"code": "PPH-003", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0627\u0633\u0648\u062f \u0645\u062e\u0631\u0632", "lastCost": 32.0, "avgCost": 32.0, "itemType": "R"}, {"code": "PPH-004", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0645\u062e\u0631\u0632 \u0641\u0636\u064a", "lastCost": 39.0, "avgCost": 43.7207771717081, "itemType": "R"}, {"code": "PPH-005", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0645\u062e\u0631\u0632 \u0634\u0641\u0627\u0641", "lastCost": 75.0, "avgCost": 55.5896038716616, "itemType": "R"}, {"code": "RUB-001", "desc": "Rubber", "lastCost": 159.0, "avgCost": 182.875, "itemType": "R"}, {"code": "DPR-001", "desc": "\u0643\u0633\u0631 \u0631\u0627\u0646\u062f\u0648\u0645", "lastCost": 38.4615, "avgCost": 38.4615, "itemType": "R"}, {"code": "PPH-007", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0645\u062e\u0631\u0632 \u0628\u064a\u062c", "lastCost": 35.0, "avgCost": 35.1771189071717, "itemType": "R"}, {"code": "PPK-002", "desc": "\u062e\u0627\u0645\u0629 \u0643\u0648\u0628\u0644\u0645\u0631 70", "lastCost": 73.0, "avgCost": 72.1088607594937, "itemType": "R"}, {"code": "M50-001", "desc": "\u062e\u0627\u0645\u0629 M50", "lastCost": 65.0, "avgCost": 74.4253393665158, "itemType": "R"}, {"code": "PET-001", "desc": "PET \u0627\u0646\u062f\u0648\u0631\u0645\u0627", "lastCost": 54.0, "avgCost": 61.9588348579682, "itemType": "R"}, {"code": "PPH-008", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0645\u062e\u0631\u0632 \u0631\u0645\u0627\u062f\u064a ", "lastCost": 37.0, "avgCost": 32.2610198789974, "itemType": "R"}, {"code": "PPK-101", "desc": "\u062e\u0627\u0645\u0629 \u0643\u0648\u0628\u0644\u0645\u0631 \u0645\u062e\u0631\u0632 ", "lastCost": 71.0, "avgCost": 51.4602533485619, "itemType": "R"}, {"code": "PPH-009", "desc": "\u0628\u0648\u0644\u064a \u0628\u0631\u0648\u0628\u0644\u064a\u0646 \u0645\u062e\u0631\u0632 \u0627\u0628\u064a\u0636 \u0633\u0646 \u0627\u0644\u0641\u064a\u0644", "lastCost": 42.0, "avgCost": 36.7623987644901, "itemType": "R"}, {"code": "PAS-001", "desc": "\u0627\u0646\u062a\u064a \u0634\u0648\u0643 \u0627\u0628\u064a\u0636", "lastCost": 48.0, "avgCost": 50.1727167246809, "itemType": "R"}, {"code": "PPS-101", "desc": "\u0628\u0648\u0644\u064a \u0627\u0633\u062a\u0631\u064a\u0646 \u0645\u062e\u0631\u0632", "lastCost": 73.0, "avgCost": 71.5607979734009, "itemType": "R"}, {"code": "PPK-103", "desc": "\u062e\u0627\u0645\u0629 \u0643\u0648\u0628\u0644\u0645\u0631 \u0645\u062e\u0631\u0632 \u0627\u062d\u0645\u0631", "lastCost": 48.0, "avgCost": 48.0745920745921, "itemType": "R"}, {"code": "SAN-001", "desc": "\u0633\u0627\u0646 \u0634\u0641\u0627\u0641", "lastCost": 110.0, "avgCost": 104.380952380952, "itemType": "R"}];

const SALES_DATA = [{"code":"TOY-2844-12","qty":-3,"amount":-214.8375},{"code":"POP-102-30","qty":-19,"amount":-756.97333},{"code":"EVO-002-12","qty":3231,"amount":127235.31831},{"code":"PTR-001-06","qty":1726,"amount":357651.76602},{"code":"ICN-206-24","qty":2712,"amount":103056},{"code":"CYL-002-60","qty":-65,"amount":936.75984},{"code":"PSH-001-01","qty":2396,"amount":520587.9232},{"code":"CYL-001-60","qty":1531,"amount":30445.16744},{"code":"PRM-103-12","qty":1055,"amount":131317.32947},{"code":"DIS-101-12","qty":422,"amount":62940.82795},{"code":"PCU-001-12","qty":1893,"amount":118074.34862},{"code":"LST-001-12","qty":97,"amount":6085.86085},{"code":"GT-11-325","qty":3250,"amount":4999.995},{"code":"PSP-001-48","qty":2389,"amount":55019.91105},{"code":"ICN-002-24","qty":1962,"amount":42845.06436},{"code":"PPR-001","qty":500,"amount":50000},{"code":"GT-12-275","qty":16500,"amount":29777.495},{"code":"ICN-008-48","qty":4704,"amount":65856},{"code":"CYL-205-12","qty":1308,"amount":202740},{"code":"PRM-002-06","qty":428,"amount":107608.50037},{"code":"GT-11-650","qty":12350,"amount":18500.001},{"code":"AQU-014-24","qty":5596,"amount":154457.02175},{"code":"EVO-101-12","qty":2836,"amount":319595.39239},{"code":"GT-14-220","qty":7040,"amount":32027.501},{"code":"CYL-104-12","qty":2775,"amount":256512.20777},{"code":"CYL-103-12","qty":2070,"amount":162125.02206},{"code":"PJS-001-12","qty":6732,"amount":123213.62858},{"code":"AQU-007-24","qty":14,"amount":2901.90197},{"code":"ICN-306-12","qty":1800,"amount":93600},{"code":"KEP-001-12","qty":8609,"amount":740164.58403},{"code":"SWT-001-24","qty":4579,"amount":195060.1339},{"code":"EVO-001-12","qty":2378,"amount":78806.99471},{"code":"ICN-108-24","qty":1944,"amount":54432},{"code":"CYL-202-12","qty":112,"amount":11790.301},{"code":"PCR-002-12","qty":520,"amount":51083.06567},{"code":"AQU-015-24","qty":-43,"amount":-86.01724},{"code":"PLT-001-12","qty":219,"amount":22553.72247},{"code":"ICN-106-24","qty":3624,"amount":90600},{"code":"ICN-207-24","qty":1632,"amount":51408},{"code":"MIX-001-01","qty":1459,"amount":95727.16131},{"code":"AQU-012-24","qty":4497,"amount":111048.50911},{"code":"DIS-104-12","qty":1998,"amount":254031.31749},{"code":"ICN-406-12","qty":2100,"amount":150150},{"code":"SQU-001-24","qty":-32,"amount":-1360},{"code":"ICN-104-24","qty":1800,"amount":42954.48},{"code":"GT-13-350","qty":10150,"amount":28777.476},{"code":"DIS-002-24","qty":2930,"amount":131160.23254},{"code":"PAP-001-12","qty":3909,"amount":92300.21286},{"code":"LUX-001-12","qty":3718,"amount":387105.59233},{"code":"DIS-102-12","qty":368,"amount":62963.25674},{"code":"AQU-005-24","qty":1838,"amount":31352.36895},{"code":"MIX-001-12","qty":2435,"amount":176300.43538},{"code":"PCR-001-06","qty":1016,"amount":170296.71463},{"code":"LOC-003-12","qty":2,"amount":42.8},{"code":"CYL-003-60","qty":566,"amount":22626.31343},{"code":"CYL-204-12","qty":492,"amount":39114},{"code":"AQU-002-24","qty":517,"amount":11496.10968},{"code":"POP-104-12","qty":3020,"amount":122067.01018},{"code":"PRM-001-06","qty":476,"amount":137688.89241},{"code":"ICN-005-48","qty":3408,"amount":54119.52},{"code":"STY-001-12","qty":24,"amount":1530},{"code":"AQU-010-24","qty":3700,"amount":92347.70483},{"code":"SMT-003-12","qty":1337,"amount":185039.54619},{"code":"POP-001-12","qty":9560,"amount":72038.54507},{"code":"AQU-009-12","qty":469,"amount":18539.04525},{"code":"KEP-001-36","qty":5940,"amount":285120},{"code":"SWT-002-24","qty":-444,"amount":-11720.34305},{"code":"PST-001-12","qty":13890,"amount":679732.89004},{"code":"AQU-003-24","qty":408,"amount":12632.89244},{"code":"ICN-001-24","qty":840,"amount":23513.29948},{"code":"GT-12-550","qty":4950,"amount":8999.991},{"code":"POP-103-12","qty":224,"amount":16968.9675},{"code":"PSH-004-01","qty":200,"amount":36000},{"code":"PAP-101-12","qty":7,"amount":503.63775},{"code":"ICN-601-12","qty":2700,"amount":237600},{"code":"AQU-001-24","qty":1066,"amount":26717.59045},{"code":"psh-002-01","qty":198,"amount":24750},{"code":"POP-003-12","qty":4541,"amount":88694.42003},{"code":"POP-002-12","qty":7370,"amount":83856.50368},{"code":"AQU-011-60","qty":32160,"amount":549936},{"code":"DIS-003-24","qty":2534,"amount":123929.65834},{"code":"SMT-001-12","qty":1539,"amount":95582.82691},{"code":"ICN-203-24","qty":2040,"amount":77520},{"code":"AQU-016-24","qty":45,"amount":2774.05496},{"code":"ICN-501-12","qty":3552,"amount":312576},{"code":"PDP-001-24","qty":1388,"amount":41156.44349},{"code":"CYL-201-12","qty":49,"amount":15167.62125},{"code":"DIS-103-12","qty":394,"amount":73415.99727},{"code":"AQU-010-60","qty":40560,"amount":601099.2},{"code":"SQU-002-24","qty":69,"amount":3098.5149},{"code":"ICN-105-24","qty":1392,"amount":38976},{"code":"POP-101-36","qty":8,"amount":263.73947},{"code":"AQU-011-24","qty":7140,"amount":154130.51856},{"code":"PLN-001-12","qty":-49,"amount":-540.38168},{"code":"CYL-203-12","qty":132,"amount":9900},{"code":"ICN-303-12","qty":1740,"amount":90480},{"code":"PTR-002-06","qty":1086,"amount":169879.46612},{"code":"PRM-102-12","qty":1026,"amount":86056.55282},{"code":"PDO-001-12","qty":6349,"amount":602515.41311},{"code":"AQU-004-24","qty":7368,"amount":97378.10803},{"code":"DIS-001-24","qty":2833,"amount":108768.04054},{"code":"PET-001","qty":1150,"amount":60662.5},{"code":"LST-101-12","qty":91,"amount":8017.9619},{"code":"PLN-002-12","qty":106,"amount":4708.26459},{"code":"BUB-001-12","qty":5187,"amount":435057.75198},{"code":"AQU-013-24","qty":6209,"amount":165032.04101},{"code":"PRM-101-12","qty":1091,"amount":73463.44716},{"code":"ICN-403-12","qty":2184,"amount":156156},{"code":"BRB-001-06","qty":9935,"amount":1117284.42018},{"code":"CYL-105-12","qty":5468,"amount":973603.11367},{"code":"SQU-003-24","qty":-32,"amount":-1125.4},{"code":"ICN-204-24","qty":1200,"amount":37800},{"code":"ICN-103-24","qty":3120,"amount":83372.16},{"code":"GT-15-100","qty":2500,"amount":22576},{"code":"PMP-001-24","qty":22165,"amount":1532465.25377},{"code":"ICN-107-24","qty":1680,"amount":35280},{"code":"PLY-001-12","qty":-256,"amount":-5509.23342},{"code":"SMT-002-12","qty":1291,"amount":132004.44287},{"code":"EVO-003-12","qty":5383,"amount":228012.39779}];

const FINISHED_GOODS = [{"code": "AQU-001-24", "desc": "زجاجة مياه اكوا هيرو 650 مل", "sp": 720.0, "conv": 24.0}, {"code": "AQU-002-24", "desc": "زجاجة مياه اكوا كيدز400 مل", "sp": 576.0, "conv": 24.0}, {"code": "AQU-003-24", "desc": "زجاجة مياه اكوا سبورت 650مل", "sp": 840.0, "conv": 24.0}, {"code": "AQU-004-24", "desc": "زجاجة مياه اكوا سكوير 350 مل", "sp": 1056.0, "conv": 48.0}, {"code": "AQU-005-24", "desc": "زجاجة مياه اكوا سكوير 500مل", "sp": 624.0, "conv": 24.0}, {"code": "AQU-007-24", "desc": "زجاجة مياه اكوا ويف 750مل", "sp": 960.0, "conv": 24.0}, {"code": "AQU-009-12", "desc": "زجاجة مياه اكوا ميجا 1250مل", "sp": 600.0, "conv": 12.0}, {"code": "AQU-009-60", "desc": "زجاجة مياه اكوا ميجا 1250مل تعبئة 60", "sp": null, "conv": 60.0}, {"code": "AQU-010-24", "desc": "زجاجة مياه اكوا سكوير 1 لتر  ", "sp": 912.0, "conv": 24.0}, {"code": "AQU-010-60", "desc": "زجاجة مياه اكوا سكوير 1لتر تعبئة 60  ", "sp": null, "conv": 60.0}, {"code": "AQU-011-24", "desc": "زجاجة مياه اكوا سكوير 1.5لتر  ", "sp": 1080.0, "conv": 24.0}, {"code": "AQU-011-60", "desc": "زجاجة مياه اكوا سكوير 1.5لتر تعبئة 60  ", "sp": null, "conv": 60.0}, {"code": "AQU-012-24", "desc": "زجاجة مياه اكوا بيراميدز 1لتر  ", "sp": 912.0, "conv": 24.0}, {"code": "AQU-013-24", "desc": "زجاجة مياه اكوا بيراميدز 1.5لتر  ", "sp": 1080.0, "conv": 24.0}, {"code": "AQU-014-24", "desc": "زجاجة مياه اكوا تانك 2.2 لتر", "sp": 1152.0, "conv": 24.0}, {"code": "AQU-015-24", "desc": "زجاجة مياه اكوا اينرجي550 مل ", "sp": 840.0, "conv": 24.0}, {"code": "AQU-016-24", "desc": "زجاجة مياه اكوا اينرجي750مل ", "sp": 750.0, "conv": 24.0}, {"code": "BRB-001-06", "desc": "خبازة بريدلي", "sp": 1164.0, "conv": 6.0}, {"code": "BUB-001-12", "desc": "علبة مناديل بابلز ", "sp": 1470.0, "conv": 12.0}, {"code": "BUB-001-60", "desc": "علبة مناديل بابلز تعبئة 60", "sp": null, "conv": 60.0}, {"code": "CYL-001-60", "desc": "علبة بهارات سيليو فرداني 300 مل ألوان", "sp": 1350.0, "conv": 60.0}, {"code": "CYL-002-60", "desc": "علبة بهارات سيليو فرداني 450 مل ألوان", "sp": 1500.0, "conv": 60.0}, {"code": "CYL-003-60", "desc": "علبة بهارات سيليو فرداني 750 مل الوان", "sp": 2250.0, "conv": 60.0}, {"code": "CYL-011-60", "desc": "علبة بهارات سيليو فرداني 300 مل خشبي", "sp": null, "conv": 60.0}, {"code": "CYL-012-60", "desc": "علبة بهارات سيليو فرداني 450 مل خشبي", "sp": null, "conv": 60.0}, {"code": "CYL-101-12", "desc": "علبة بهارات سيليو 9 في 1 ألوان", "sp": null, "conv": 12.0}, {"code": "CYL-102-12", "desc": "علبة بهارات سيليو 3 في 1 ألوان", "sp": null, "conv": 12.0}, {"code": "CYL-103-12", "desc": "ستاند رباعي سيليو 300 مل الوان", "sp": 1350.0, "conv": 12.0}, {"code": "CYL-104-12", "desc": "ستاند رباعي سيليو 450 مل الوان", "sp": 1575.0, "conv": 12.0}, {"code": "CYL-105-12", "desc": "ستاند بهارات  سيليو 2 دور ألوان", "sp": 3000.0, "conv": 12.0}, {"code": "CYL-201-12", "desc": "علبة بهارات سيليو 9 في 1 خشبي", "sp": 3375.0, "conv": 12.0}, {"code": "CYL-202-12", "desc": "علبة بهارات سيليو 3 في 1 خشبي", "sp": 1200.0, "conv": 12.0}, {"code": "CYL-203-12", "desc": "ستاند رباعي سيليو 300 مل خشبي", "sp": null, "conv": 12.0}, {"code": "CYL-204-12", "desc": "ستاند رباعي سيليو 450 مل خشبي", "sp": null, "conv": 12.0}, {"code": "CYL-205-12", "desc": "ستاند بهارات  سيليو 2 دور خشبي", "sp": null, "conv": 12.0}, {"code": "CYL-305-12", "desc": "ستاند بهارات  سيليو 2 دور ابيض", "sp": null, "conv": 12.0}, {"code": "DIS-001-24", "desc": "علبة ديسكفري مفرد صغير 800", "sp": 1353.0, "conv": 24.0}, {"code": "DIS-002-24", "desc": "علبة ديسكفري مفرد وسط 1000", "sp": 1585.0, "conv": 24.0}, {"code": "DIS-003-24", "desc": "علبة ديسكفري مفرد كبير 1300", "sp": 1813.0, "conv": 24.0}, {"code": "DIS-101-12", "desc": "علبة ديسكفري ثلاثي  صغير800", "sp": 1981.0, "conv": 12.0}, {"code": "DIS-102-12", "desc": "علبة ديسكفري ثلاثي وسط 1000", "sp": 2313.0, "conv": 12.0}, {"code": "DIS-103-12", "desc": "علبة ديسكفري ثلاثي كبير 1300", "sp": 2638.0, "conv": 12.0}, {"code": "DIS-104-12", "desc": "علبة ديسكفري ثلاثي مشكل", "sp": 2250.0, "conv": 12.0}, {"code": "DPH-001", "desc": "كسر بولي بروبلين", "sp": null, "conv": null}, {"code": "DPS-001-01", "desc": "كسر كريستال ", "sp": null, "conv": null}, {"code": "EVO-001-12", "desc": "علبة تخزين ايفو صغيرة 1.25 لتر", "sp": 494.0, "conv": 12.0}, {"code": "EVO-001-24", "desc": "علبة تخزين ايفو صغيرة 1.25 لتر 24 قطعة", "sp": 530.0, "conv": 24.0}, {"code": "EVO-002-12", "desc": "علبة تخزين ايفو وسط 2 لتر", "sp": 569.0, "conv": 12.0}, {"code": "EVO-002-24", "desc": "علبة تخزين ايفو وسط 2 لتر 24 قطعة", "sp": 610.0, "conv": 24.0}, {"code": "EVO-003-12", "desc": "علبة تخزين ايفو كبير 2.75 لتر", "sp": 663.0, "conv": 12.0}, {"code": "EVO-003-24", "desc": "علبة تخزين ايفو كبير 2.75 لتر 24 قطعة", "sp": 710.0, "conv": 24.0}, {"code": "EVO-101-12", "desc": "طقم تخزين ايفو مشكل", "sp": 1688.0, "conv": 12.0}, {"code": "GLD-001-650", "desc": "علبة جولدن مقاس 0.20 لتر ", "sp": null, "conv": null}, {"code": "GLD-002-275", "desc": "علبة  جولدن مقاس 0.3 لتر تعبئة 275 ", "sp": null, "conv": 550.0}, {"code": "GLD-002-550", "desc": "علبة  جولدن مقاس 0.3 لتر ", "sp": null, "conv": 550.0}, {"code": "GLD-003-350", "desc": "علبة جولدن مقاس 0.5لتر ", "sp": null, "conv": 350.0}, {"code": "GLD-004-220", "desc": "علبة جولدن مقاس 1لتر ", "sp": null, "conv": 220.0}, {"code": "GLD-005-100", "desc": "علبة جولدن مقاس 2لتر ", "sp": null, "conv": 100.0}, {"code": "GT-001-650", "desc": "علبة مقاس 0.20 لتر ", "sp": 950.0, "conv": 650.0}, {"code": "GT-02-550", "desc": "علبة مقاس 0.3 لتر ", "sp": 950.0, "conv": 550.0}, {"code": "GT-03-350", "desc": "علبة مقاس 0.5لتر ", "sp": 950.0, "conv": 350.0}, {"code": "GT-04-220", "desc": "علبة مقاس 1لتر ", "sp": 950.0, "conv": 220.0}, {"code": "GT-05-100", "desc": "علبة مقاس 2لتر ", "sp": 850.0, "conv": 100.0}, {"code": "GT-11-325", "desc": "علبة مقاس 0.20 لتر لون ابيض عدد 325 ", "sp": null, "conv": 650.0}, {"code": "GT-11-650", "desc": "علبة مقاس 0.20 لتر لون ابيض ", "sp": 950.0, "conv": 650.0}, {"code": "GT-12-275", "desc": "علبة مقاس 0.3 لتر لون ابيض عدد 275", "sp": null, "conv": 550.0}, {"code": "GT-12-550", "desc": "علبة مقاس 0.3 لتر لون ابيض", "sp": 950.0, "conv": 550.0}, {"code": "GT-13-350", "desc": "علبة مقاس 0.5لتر لون ابيض ", "sp": 950.0, "conv": 350.0}, {"code": "GT-14-220", "desc": "علبة مقاس 1لتر لون ابيض ", "sp": 950.0, "conv": 220.0}, {"code": "GT-15-100", "desc": "علبة مقاس 2لتر لون ابيض ", "sp": 850.0, "conv": 100.0}, {"code": "I01-2601107", "desc": "جردل تربو ستار 13 لتر", "sp": null, "conv": null}, {"code": "I01-2601109", "desc": "جردل ايكو تريندي 13 لتر", "sp": null, "conv": null}, {"code": "I01-2601122", "desc": "جردل موب سيت 15 لتر", "sp": null, "conv": null}, {"code": "I01-2601123", "desc": "جردل ستايل باركيه", "sp": null, "conv": null}, {"code": "I01-2601124", "desc": "جردل سليم باركيه", "sp": null, "conv": null}, {"code": "I01-2601125", "desc": "جردل فوكس باركيه", "sp": null, "conv": null}, {"code": "I01-2601137", "desc": "جردل ستاركس بالبدال", "sp": null, "conv": null}, {"code": "I01-2601206", "desc": "مساحة تريو", "sp": null, "conv": null}, {"code": "I01-2601208", "desc": "مساحة مايكروفايبر", "sp": null, "conv": null}, {"code": "I01-2601217", "desc": "جاروف بالعصاية", "sp": null, "conv": null}, {"code": "I01-2601222", "desc": "شرشوبة مايكرو فايبر سوبر", "sp": null, "conv": null}, {"code": "I01-2601223", "desc": "شرشوبة مايكروفايبر", "sp": null, "conv": null}, {"code": "I01-2601629", "desc": "سبت غسيل مع ارجل", "sp": null, "conv": null}, {"code": "I01-2601632", "desc": "طقم حمام 5 قطع", "sp": null, "conv": null}, {"code": "I01-2602100", "desc": "جردل اوتامتيك 35 لتر", "sp": null, "conv": null}, {"code": "I02-2210", "desc": "سلة مهملات 50 لتر بالغطاء", "sp": null, "conv": null}, {"code": "I02-2228", "desc": "زبالة رخامى 6.5 لتر", "sp": null, "conv": null}, {"code": "I02-2229", "desc": "زبالة رخامى13 لتر", "sp": null, "conv": null}, {"code": "I02-2230", "desc": "زبالة رخامى20لتر", "sp": null, "conv": null}, {"code": "I02-390", "desc": "سلة مهملات 50 لتر ازمير مروحة", "sp": null, "conv": null}, {"code": "I02-495", "desc": "بوتى أطفال ", "sp": null, "conv": null}, {"code": "I02-520", "desc": "شنطة أطفال كوكى", "sp": null, "conv": null}, {"code": "I02-601", "desc": " طقم حمام راتان 5 قطع", "sp": null, "conv": null}, {"code": "I02-605", "desc": " سبت غسيل راتان 60 لتر بيد خشب", "sp": null, "conv": null}, {"code": "I02-606", "desc": " سلة غسيل راتان 30 لتر بيد خشب", "sp": null, "conv": null}, {"code": "I02-611", "desc": " سلة غسيل", "sp": null, "conv": null}, {"code": "I02-618", "desc": " فرشاة حمام راتان", "sp": null, "conv": null}, {"code": "I02-619", "desc": " سلة مهملات 20 لتر بالدواسة", "sp": null, "conv": null}, {"code": "I02-622", "desc": " سلة مشابك 7لتر بيد خشب", "sp": null, "conv": null}, {"code": "I02-625", "desc": " سلة مهملات راتان 5.5 لتر", "sp": null, "conv": null}, {"code": "ICN-001-24", "desc": "برطمان ايكون 1400 مل", "sp": 705.0, "conv": 24.0}, {"code": "ICN-002-24", "desc": "برطمان ايكون 800 مل", "sp": 585.0, "conv": 24.0}, {"code": "ICN-005-48", "desc": "برطمان ايكون 8 1000مل ", "sp": 1680.0, "conv": 48.0}, {"code": "ICN-008-48", "desc": "برطمان ايكون سكوير 1300مل ", "sp": null, "conv": 48.0}, {"code": "ICN-103-24", "desc": "برطمان ايكون 8 400 مل 4*1", "sp": 1440.0, "conv": 24.0}, {"code": "ICN-104-24", "desc": "برطمان ايكون 8 750 مل 2*1 ", "sp": 1320.0, "conv": 24.0}, {"code": "ICN-105-24", "desc": "برطمان ايكون 8 1000مل 2*1 ", "sp": null, "conv": 24.0}, {"code": "ICN-106-24", "desc": "برطمان ايكون سكوير 400 مل 4*1", "sp": null, "conv": 24.0}, {"code": "ICN-107-24", "desc": "برطمان ايكون سكوير 800مل 2*1 ", "sp": null, "conv": 24.0}, {"code": "ICN-108-24", "desc": "برطمان ايكون سكوير 1300مل 2*1 ", "sp": null, "conv": 24.0}, {"code": "ICN-203-24", "desc": "برطمان ايكون 8 400 مل 6*1", "sp": null, "conv": 24.0}, {"code": "ICN-204-24", "desc": "برطمان ايكون 8 750 مل 3*1", "sp": null, "conv": 24.0}, {"code": "ICN-206-24", "desc": "برطمان ايكون سكوير 400 مل 6*1", "sp": null, "conv": 24.0}, {"code": "ICN-207-24", "desc": "برطمان ايكون سكوير 800 مل 3*1", "sp": null, "conv": 24.0}, {"code": "ICN-303-12", "desc": "برطمان ايكون 8 400 مل 8*1", "sp": null, "conv": 12.0}, {"code": "ICN-306-12", "desc": "برطمان ايكون سكوير 400 مل 8*1", "sp": null, "conv": 12.0}, {"code": "ICN-403-12", "desc": "برطمان ايكون 8 400 مل 12*1", "sp": 2100.0, "conv": 12.0}, {"code": "ICN-406-12", "desc": "برطمان ايكون سكوير 400 مل 12*1", "sp": null, "conv": 12.0}, {"code": "ICN-501-12", "desc": "ستاند مرطبان ايكون مربع7*1", "sp": 2580.0, "conv": 12.0}, {"code": "ICN-601-12", "desc": "ستاند مرطبان ايكون مضلع 7*1", "sp": null, "conv": 12.0}, {"code": "KEP-001-12", "desc": "حافظة مخبوزات كيبير", "sp": 1375.0, "conv": 12.0}, {"code": "KEP-001-36", "desc": "حافظة مخبوزات كيبير عدد 36", "sp": null, "conv": 36.0}, {"code": "L01-001-100", "desc": " مشبك 24 قطعة عريض", "sp": null, "conv": null}, {"code": "LOC-001-12", "desc": "جردل توينز الهلال والنجمة الذهبية", "sp": null, "conv": null}, {"code": "LOC-002-12", "desc": "جردل بیضاوي بالعصارة الهلال", "sp": null, "conv": null}, {"code": "LOC-003-12", "desc": "فرشة هلين اسود الهلال - الذهبية", "sp": null, "conv": null}, {"code": "LOC-004-12", "desc": "فرشة روما لوكس - الهلال", "sp": null, "conv": null}, {"code": "LOC-005-12", "desc": "فرشة سجاد وموكيت روتو الهلال - الذهبية", "sp": null, "conv": null}, {"code": "LOC-006-12", "desc": "زعافة كورة بالاريال مدهون", "sp": null, "conv": null}, {"code": "LOC-007-12", "desc": "فرشة بالجاروف توام - الذهبية", "sp": null, "conv": null}, {"code": "LOC-008-12", "desc": "فرشة ايديال ناعمة الهلال", "sp": null, "conv": null}, {"code": "LOC-009-12", "desc": "فرشة ايديال خشنة - الهلال", "sp": null, "conv": null}, {"code": "LOC-010-12", "desc": "فرشة بلاط باليد هلال", "sp": null, "conv": null}, {"code": "LOC-011-12", "desc": "فرشة تواليت بالقاعدة بامبو - الهلال", "sp": null, "conv": null}, {"code": "LOC-012-12", "desc": "مساحة هلال صغيرة 40 سم ", "sp": null, "conv": null}, {"code": "LOC-013-12", "desc": "مساحة هلال كبيرة 46 سم ", "sp": null, "conv": null}, {"code": "LOC-014-12", "desc": "فرشة فور جى الوان ", "sp": null, "conv": null}, {"code": "LOC-015-12", "desc": "فرشة جولدن ماكس جرانيت ", "sp": null, "conv": null}, {"code": "LOC-016-12", "desc": "مساحة سوبر ستار المصرية ", "sp": null, "conv": null}, {"code": "LOC-017-12", "desc": "جاروف هلال بالكاوتشة كلاسيك ", "sp": null, "conv": null}, {"code": "LOC-018-12", "desc": "فرشة تواليت عادة الهلال", "sp": null, "conv": null}, {"code": "LOC-019-12", "desc": "يد زان صینی", "sp": null, "conv": null}, {"code": "LST-001-12", "desc": "لانش بوكس سمارت", "sp": 906.0, "conv": 12.0}, {"code": "LST-101-12", "desc": "مجموعة لانش بوكس سمارت", "sp": 1125.0, "conv": 12.0}, {"code": "LUX-001-06", "desc": "رف لوكس 6 قطع", "sp": 500.0, "conv": 6.0}, {"code": "LUX-001-12", "desc": "رف لوكس", "sp": 1725.0, "conv": 12.0}, {"code": "MB-013", "desc": "صباغ أبيض نقطة", "sp": null, "conv": null}, {"code": "MB-014", "desc": "صباغ بني فاتح ", "sp": null, "conv": null}, {"code": "MB-016", "desc": "صباغ أحمر تركي", "sp": null, "conv": null}, {"code": "MB-017", "desc": "صباغ أخضر تركي", "sp": null, "conv": null}, {"code": "MB-021", "desc": "صباغ فيروزي غامق", "sp": null, "conv": null}, {"code": "MB-025", "desc": "صباغ نهدي تركي", "sp": null, "conv": null}, {"code": "MB-026", "desc": "صباغ بيج مونس سويدي", "sp": null, "conv": null}, {"code": "MB-027", "desc": "صباغ فبروزي تركي", "sp": null, "conv": null}, {"code": "MB-028", "desc": "صباغ سماوي تركي (لبني)", "sp": null, "conv": null}, {"code": "MB-030", "desc": "صباغ فضي", "sp": null, "conv": null}, {"code": "MB-032", "desc": "صباغ موف", "sp": null, "conv": null}, {"code": "MB-035", "desc": "صباغ اصفر ميتالك", "sp": null, "conv": null}, {"code": "MB-036", "desc": "صباغ احمر ميتالك", "sp": null, "conv": null}, {"code": "MB-037", "desc": "صباغ اسود", "sp": null, "conv": null}, {"code": "MB-038", "desc": "صباغ ابيض منمش", "sp": null, "conv": null}, {"code": "MB-039", "desc": "صباغ ابيض ", "sp": null, "conv": null}, {"code": "MB-040", "desc": "صباغ كشميري", "sp": null, "conv": null}, {"code": "MB-041", "desc": "صباغ كحلي ميتالك", "sp": null, "conv": null}, {"code": "MB-043", "desc": "صباغ ازرق بيبسي", "sp": null, "conv": null}, {"code": "MB-044", "desc": "صباغ اخضر ميتالك", "sp": null, "conv": null}, {"code": "MB-055", "desc": "صباغ زهري ميتالك", "sp": null, "conv": null}, {"code": "MIX-001-01", "desc": "دوراق ميكسر بالكرتونة", "sp": 840.0, "conv": 12.0}, {"code": "MIX-001-12", "desc": "دوراق ميكسر", "sp": 1050.0, "conv": 12.0}, {"code": "PAP-001-12", "desc": "طبق مقبلات و تسالي/5 عين", "sp": 1594.0, "conv": 60.0}, {"code": "PAP-101- 12", "desc": "طقم طبق تسالي عدد 12 طقم ", "sp": null, "conv": 12.0}, {"code": "PAP-101-12", "desc": "طقم طبق مقبلات و تسالي ثلاثي", "sp": 849.5, "conv": 18.0}, {"code": "PCR-001-06", "desc": "ركنة 4 دور", "sp": 1281.0, "conv": 6.0}, {"code": "PCR-002-12", "desc": "ركنة 2 دور", "sp": 1400.0, "conv": 12.0}, {"code": "PCR-003-08", "desc": "ركنة 3 دور", "sp": null, "conv": null}, {"code": "PCU-001-12", "desc": "قطاعة خضار", "sp": 950.0, "conv": 12.0}, {"code": "PDO-001-12", "desc": "منظم اطباق بيرفكت", "sp": 1650.0, "conv": 12.0}, {"code": "PDP-001-24", "desc": "طبق مضلع", "sp": 725.0, "conv": 24.0}, {"code": "PET-001", "desc": "PET اندورما", "sp": null, "conv": null}, {"code": "PJS-001-12", "desc": "عصارة موالح", "sp": 294.0, "conv": 12.0}, {"code": "PLN-001-12", "desc": "لانش بوكس كيدز", "sp": 263.0, "conv": 12.0}, {"code": "PLN-002-12", "desc": "لانش بوكس هيرو", "sp": 563.0, "conv": 12.0}, {"code": "PLT-001-12", "desc": "طبق بلايت بلس ", "sp": 1350.0, "conv": 12.0}, {"code": "PLY-001-12", "desc": "لانش بوكس بلاى", "sp": 413.0, "conv": 12.0}, {"code": "PMP-001-24", "desc": "منظم معالق بيرفكت ", "sp": 2250.0, "conv": 24.0}, {"code": "POL-001-00", "desc": "علبة بوليش كاملة", "sp": null, "conv": 1.0}, {"code": "POP-001-12", "desc": "طبق اوفال صغير", "sp": 1344.0, "conv": 144.0}, {"code": "POP-002-12", "desc": "طبق اوفال وسط", "sp": 1550.0, "conv": 108.0}, {"code": "POP-003-12", "desc": "طبق اوفال كبير", "sp": 1125.0, "conv": 48.0}, {"code": "POP-101-12", "desc": "طقم طبق اوفال ثلاثي صغير", "sp": 762.0, "conv": 48.0}, {"code": "POP-101-36", "desc": "طقم طبق اوفال ثلاثي صغير عدد 36 طقم", "sp": null, "conv": 36.0}, {"code": "POP-102-12", "desc": "طقم طبق اوفال ثلاثي وسط", "sp": 843.0, "conv": 36.0}, {"code": "POP-102-30", "desc": "طقم طبق اوفال ثلاثي وسط عدد 30 طقم ", "sp": null, "conv": 30.0}, {"code": "POP-103- 12", "desc": "طقم طبق اوفال ثلاثي كبير عدد 12 طقم ", "sp": null, "conv": 12.0}, {"code": "POP-103-12", "desc": "طقم طبق اوفال ثلاثي كبير", "sp": 515.5, "conv": 16.0}, {"code": "POP-104-12", "desc": "طقم طبق اوفال ثلاثي مشكل", "sp": 1100.0, "conv": 24.0}, {"code": "PPH-001", "desc": "بولي بروبلين H26ML", "sp": null, "conv": null}, {"code": "PPH-004", "desc": "بولي بروبلين مخرز فضي", "sp": null, "conv": null}, {"code": "PPH-005", "desc": "بولي بروبلين مخرز شفاف", "sp": null, "conv": null}, {"code": "PPH-007", "desc": "بولي بروبلين مخرز بيج", "sp": null, "conv": null}, {"code": "PPH-008", "desc": "بولي بروبلين مخرز رمادي ", "sp": null, "conv": null}, {"code": "PPH-009", "desc": "بولي بروبلين مخرز ابيض سن الفيل", "sp": null, "conv": null}, {"code": "PPK-001", "desc": "خامة كوبلمر", "sp": null, "conv": null}, {"code": "PPK-002", "desc": "خامة كوبلمر 70", "sp": null, "conv": null}, {"code": "PPK-101", "desc": "خامة كوبلمر مخرز ", "sp": null, "conv": null}, {"code": "PPR-001", "desc": "بولي بروبلين راندوم R40", "sp": null, "conv": null}, {"code": "PPS-001", "desc": "بولي استرين", "sp": null, "conv": null}, {"code": "PRM-001-06", "desc": "علبة توابل 150 مل", "sp": 2138.0, "conv": 6.0}, {"code": "PRM-002-06", "desc": "علبة توابل 300 مل", "sp": 1950.0, "conv": 6.0}, {"code": "PRM-002-550", "desc": "علبة  برايم جي تي مقاس 0.3 لتر ", "sp": null, "conv": null}, {"code": "PRM-003-350", "desc": "علبة برايم جي تي مقاس 0.5لتر ", "sp": null, "conv": null}, {"code": "PRM-004-220", "desc": "علبة برايم جي تي مقاس 1لتر ", "sp": null, "conv": null}, {"code": "PRM-005-100", "desc": "علبة برايم جي تي مقاس 2لتر ", "sp": null, "conv": null}, {"code": "PRM-101-12", "desc": "ستاند توابل برايم رباعي 150 مل", "sp": 863.0, "conv": 12.0}, {"code": "PRM-102-12", "desc": "ستاند توابل برايم رباعي 300 مل", "sp": 1038.0, "conv": 12.0}, {"code": "PRM-103-12", "desc": "ستاند توابل برايم ثماني مشكل", "sp": 1706.0, "conv": 12.0}, {"code": "PSH-001-01", "desc": "جزامة 5 دور بيرفكت", "sp": 275.0, "conv": 1.0}, {"code": "psh-002-01", "desc": "جزامة 4دور بيرفكت", "sp": null, "conv": 1.0}, {"code": "PSH-003-01", "desc": "جزامة 3دور بيرفكت", "sp": null, "conv": null}, {"code": "PSH-004-01", "desc": "جزامة 6 دور بيرفكت", "sp": null, "conv": 1.0}, {"code": "PSP-001-48", "desc": "طبق سرفيس", "sp": 1219.0, "conv": 48.0}, {"code": "PST-001-12", "desc": "مصفى 2 في 1", "sp": 813.0, "conv": 12.0}, {"code": "PST-002-200", "desc": "مصفى", "sp": null, "conv": null}, {"code": "PTR-001-06", "desc": "ترولي بيرفكت 4 دور", "sp": 1625.0, "conv": 6.0}, {"code": "PTR-002-06", "desc": "ترولي بيرفكت 3دور", "sp": 1219.0, "conv": 6.0}, {"code": "RUB-001", "desc": "Rubber", "sp": null, "conv": null}, {"code": "SHV-001-1500", "desc": "مشحاف جولدين لون احمر", "sp": null, "conv": 1500.0}, {"code": "SMT-001-12", "desc": "حافظة طعام سمارت فرداني", "sp": 858.0, "conv": 12.0}, {"code": "SMT-002-12", "desc": "حافظة طعام سمارت ثنائي", "sp": 1413.0, "conv": 12.0}, {"code": "SMT-003-12", "desc": "حافظة طعام سمارت ثلاثي", "sp": 1875.0, "conv": 12.0}, {"code": "SQU-001-24", "desc": "مجموعه إزازة مياة ولانش بوكس هيرو", "sp": 1500.0, "conv": 24.0}, {"code": "SQU-002-24", "desc": "مجموعه إزازة مياة ولانش بوكس كيدز", "sp": 930.0, "conv": 24.0}, {"code": "SQU-003-24", "desc": "مجموعه إزازة مياة ولانش بوكس بلاى", "sp": 1241.0, "conv": 24.0}, {"code": "STY-001-12", "desc": "ستاند فرشاة استايل", "sp": null, "conv": null}, {"code": "SWT-001-24", "desc": "قالب حلويات صغير", "sp": 1350.0, "conv": 24.0}, {"code": "SWT-002-24", "desc": "قالب تحضير متعدد الاستخدامات", "sp": 1200.0, "conv": 24.0}, {"code": "TOY-02929-12", "desc": "طبلة كبيرة ", "sp": null, "conv": null}, {"code": "TOY-02936-24", "desc": "بندقية كلاشنكوف على كرت", "sp": null, "conv": null}, {"code": "TOY-02981-24", "desc": "مسدس جيمس بوند على كرت", "sp": null, "conv": null}, {"code": "TOY-2097-12", "desc": "يخت لوكس", "sp": null, "conv": null}, {"code": "TOY-2141-12", "desc": "زورق بحر", "sp": null, "conv": null}, {"code": "TOY-2325-12", "desc": "يخت بحر", "sp": null, "conv": null}, {"code": "TOY-2332-12", "desc": "براويطة بحر", "sp": null, "conv": null}, {"code": "TOY-2479-24", "desc": "مسدس طبنجة على كرت ", "sp": null, "conv": null}, {"code": "TOY-2486-12", "desc": "سيارة تريلا  95 ", "sp": null, "conv": null}, {"code": "TOY-2592-12", "desc": "سيارة جيب وينجر", "sp": null, "conv": null}, {"code": "TOY-2646-36", "desc": "سيارة مينى كوبر", "sp": null, "conv": null}, {"code": "TOY-2653-36", "desc": "سيارة بنزين 96", "sp": null, "conv": null}, {"code": "TOY-2684-12", "desc": "الدب السعيد", "sp": null, "conv": null}, {"code": "TOY-2707-12", "desc": "بوسيت مطبخ", "sp": null, "conv": null}, {"code": "TOY-2714-12", "desc": "بوسيت شاى", "sp": null, "conv": null}, {"code": "TOY-2721-12", "desc": "صنية دورا او مطبخ دورا", "sp": null, "conv": null}, {"code": "TOY-2738-12", "desc": "عربية سوبر ماركت", "sp": null, "conv": null}, {"code": "TOY-2745-12", "desc": " كرسى بيبى", "sp": null, "conv": null}, {"code": "TOY-2769-12", "desc": "حقيبة هيرو مكعبات", "sp": null, "conv": null}, {"code": "TOY-2776-12", "desc": "قطار هيرو", "sp": null, "conv": null}, {"code": "TOY-2783-12", "desc": "سيارة فرارى", "sp": null, "conv": null}, {"code": "TOY-2790-12", "desc": "سيارة هيرو مكعبات", "sp": null, "conv": null}, {"code": "TOY-2806-12", "desc": "شنطة مطبخ سالى ابو هند رستم", "sp": null, "conv": null}, {"code": "TOY-2813-24", "desc": "بندقية قطز", "sp": null, "conv": null}, {"code": "TOY-2820-24", "desc": "مسدس عز الدين ", "sp": null, "conv": null}, {"code": "TOY-2844-12", "desc": "مطبخ 3 دور", "sp": null, "conv": null}, {"code": "TOY-2851-12", "desc": "مطبخ 2 دور", "sp": null, "conv": null}, {"code": "TOY-2868-24", "desc": "بندقية كارت صوت", "sp": null, "conv": null}, {"code": "TOY-2875-12", "desc": "سيارة بيتش باجى", "sp": null, "conv": null}, {"code": "TOY-2882-12", "desc": "سيارة صاروخ او طيارة", "sp": null, "conv": null}, {"code": "TOY-2899-12", "desc": "كورة كبيرة", "sp": null, "conv": null}, {"code": "TOY-2905-12", "desc": "كورة صغيرة", "sp": null, "conv": null}, {"code": "TOY-2912-24", "desc": "بطة ويكى او بطة عجلة", "sp": null, "conv": null}, {"code": "TOY-2929-12", "desc": "طبلة رمضان", "sp": null, "conv": null}, {"code": "TOY-2936-24", "desc": "بندقية كلاشنكوف", "sp": null, "conv": null}, {"code": "TOY-2943-12", "desc": " أتوبيس المكعبات", "sp": null, "conv": null}, {"code": "TOY-2950-12", "desc": " سيارة فوكس او سيارة فاتن حمامة", "sp": null, "conv": null}, {"code": "TOY-2967-12", "desc": " بيت المكعبات", "sp": null, "conv": null}, {"code": "TOY-2974-12", "desc": "  مطبخ فرحة", "sp": null, "conv": null}, {"code": "TOY-2981-24", "desc": "  مسدس جيمس بوند", "sp": null, "conv": null}, {"code": "TOY-2998-24", "desc": "  سيارة فرمولا", "sp": null, "conv": null}];

const INITIAL_DATA = [{"code": "AQU-001-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0647\u064a\u0631\u0648 650 \u0645\u0644", "conv": 24, "sp": 720.0, "rawBoxCost": 139.2982, "boxCost": 188.0526, "profit": 531.9474, "profitPct": 73.88, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-001-24", "qty": 1.0, "unitCost": 8.0, "lineCost": 8.0}, {"child": "M50-001", "qty": 0.12, "unitCost": 74.4253393665158, "lineCost": 8.931}, {"child": "MB-032", "qty": 0.00204, "unitCost": 190.363636363636, "lineCost": 0.3883}, {"child": "MB-035", "qty": 0.00204, "unitCost": 186.397467572576, "lineCost": 0.3803}, {"child": "MB-036", "qty": 0.00204, "unitCost": 188.125, "lineCost": 0.3838}, {"child": "MB-043", "qty": 0.00204, "unitCost": 170.799180327869, "lineCost": 0.3484}, {"child": "MB-044", "qty": 0.00204, "unitCost": 188.547619047619, "lineCost": 0.3846}, {"child": "MB-055", "qty": 0.00204, "unitCost": 214.316753926702, "lineCost": 0.4372}, {"child": "MB-201", "qty": 0.0012, "unitCost": 323.888888888889, "lineCost": 0.3887}, {"child": "MB-202", "qty": 0.0012, "unitCost": 353.333333333333, "lineCost": 0.424}, {"child": "MB-203", "qty": 0.0012, "unitCost": 344.815950920245, "lineCost": 0.4138}, {"child": "MB-205", "qty": 0.0012, "unitCost": 297.352941176471, "lineCost": 0.3568}, {"child": "MB-206", "qty": 0.0012, "unitCost": 415.0, "lineCost": 0.498}, {"child": "MB-207", "qty": 0.0012, "unitCost": 415.0, "lineCost": 0.498}, {"child": "PET-001", "qty": 1.44, "unitCost": 61.9588348579682, "lineCost": 89.2207}, {"child": "PPH-001", "qty": 0.312, "unitCost": 67.4507739938081, "lineCost": 21.0446}, {"child": "ST-AQU-001", "qty": 24.0, "unitCost": 0.3, "lineCost": 7.2}]}, {"code": "AQU-002-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0643\u064a\u062f\u0632400 \u0645\u0644", "conv": 24, "sp": 576.0, "rawBoxCost": 129.1356, "boxCost": 174.3331, "profit": 401.6669, "profitPct": 69.73, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-002-24", "qty": 1.0, "unitCost": 7.52776805251641, "lineCost": 7.5278}, {"child": "M50-001", "qty": 0.12, "unitCost": 74.4253393665158, "lineCost": 8.931}, {"child": "MB-032", "qty": 0.00204, "unitCost": 190.363636363636, "lineCost": 0.3883}, {"child": "MB-035", "qty": 0.00204, "unitCost": 186.397467572576, "lineCost": 0.3803}, {"child": "MB-036", "qty": 0.00204, "unitCost": 188.125, "lineCost": 0.3838}, {"child": "MB-043", "qty": 0.00204, "unitCost": 170.799180327869, "lineCost": 0.3484}, {"child": "MB-044", "qty": 0.00204, "unitCost": 188.547619047619, "lineCost": 0.3846}, {"child": "MB-055", "qty": 0.00204, "unitCost": 214.316753926702, "lineCost": 0.4372}, {"child": "MB-201", "qty": 0.0016, "unitCost": 323.888888888889, "lineCost": 0.5182}, {"child": "MB-202", "qty": 0.0016, "unitCost": 353.333333333333, "lineCost": 0.5653}, {"child": "MB-203", "qty": 0.0016, "unitCost": 344.815950920245, "lineCost": 0.5517}, {"child": "MB-205", "qty": 0.0016, "unitCost": 297.352941176471, "lineCost": 0.4758}, {"child": "MB-206", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "MB-207", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "PET-001", "qty": 1.2, "unitCost": 61.9588348579682, "lineCost": 74.3506}, {"child": "PPH-001", "qty": 0.312, "unitCost": 67.4507739938081, "lineCost": 21.0446}, {"child": "ST-AQU-002", "qty": 24.0, "unitCost": 0.48, "lineCost": 11.52}]}, {"code": "AQU-003-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0628\u0648\u0631\u062a 650\u0645\u0644", "conv": 24, "sp": 840.0, "rawBoxCost": 180.8152, "boxCost": 244.1005, "profit": 595.8995, "profitPct": 70.94, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-003-24", "qty": 1.0, "unitCost": 9.63, "lineCost": 9.63}, {"child": "MB-032", "qty": 0.00456, "unitCost": 190.363636363636, "lineCost": 0.8681}, {"child": "MB-036", "qty": 0.00456, "unitCost": 188.125, "lineCost": 0.8579}, {"child": "MB-037", "qty": 0.00456, "unitCost": 109.956331877729, "lineCost": 0.5014}, {"child": "MB-043", "qty": 0.00456, "unitCost": 170.799180327869, "lineCost": 0.7788}, {"child": "MB-044", "qty": 0.00456, "unitCost": 188.547619047619, "lineCost": 0.8598}, {"child": "MB-055", "qty": 0.00456, "unitCost": 214.316753926702, "lineCost": 0.9773}, {"child": "MB-208", "qty": 0.0036, "unitCost": 241.666666666667, "lineCost": 0.87}, {"child": "PET-001", "qty": 1.44, "unitCost": 61.9588348579682, "lineCost": 89.2207}, {"child": "PPH-001", "qty": 0.912, "unitCost": 67.4507739938081, "lineCost": 61.5151}, {"child": "RUB-001", "qty": 0.036, "unitCost": 182.875, "lineCost": 6.5835}, {"child": "ST-AQU-003-1", "qty": 24.0, "unitCost": 0.339692307692308, "lineCost": 8.1526}]}, {"code": "AQU-004-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 350 \u0645\u0644", "conv": 48, "sp": 1056.0, "rawBoxCost": 231.9839, "boxCost": 313.1783, "profit": 742.8217, "profitPct": 70.34, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-004-48", "qty": 1.0, "unitCost": 12.466930545183, "lineCost": 12.4669}, {"child": "MB-032", "qty": 0.00144, "unitCost": 190.363636363636, "lineCost": 0.2741}, {"child": "MB-035", "qty": 0.00288, "unitCost": 186.397467572576, "lineCost": 0.5368}, {"child": "MB-036", "qty": 0.00144, "unitCost": 188.125, "lineCost": 0.2709}, {"child": "MB-043", "qty": 0.00144, "unitCost": 170.799180327869, "lineCost": 0.246}, {"child": "MB-044", "qty": 0.00144, "unitCost": 188.547619047619, "lineCost": 0.2715}, {"child": "MB-046", "qty": 0.00288, "unitCost": 112.599364069952, "lineCost": 0.3243}, {"child": "MB-047", "qty": 0.00144, "unitCost": 114.0, "lineCost": 0.1642}, {"child": "MB-048", "qty": 0.00144, "unitCost": 110.0, "lineCost": 0.1584}, {"child": "MB-049", "qty": 0.00144, "unitCost": 110.0, "lineCost": 0.1584}, {"child": "MB-050", "qty": 0.00144, "unitCost": 114.0, "lineCost": 0.1642}, {"child": "MB-051", "qty": 0.00144, "unitCost": 111.527777777778, "lineCost": 0.1606}, {"child": "MB-055", "qty": 0.00144, "unitCost": 214.316753926702, "lineCost": 0.3086}, {"child": "PET-001", "qty": 2.4, "unitCost": 61.9588348579682, "lineCost": 148.7012}, {"child": "PPH-001", "qty": 0.672, "unitCost": 67.4507739938081, "lineCost": 45.3269}, {"child": "ST-AQU-004-1", "qty": 48.0, "unitCost": 0.467728055077453, "lineCost": 22.4509}]}, {"code": "AQU-005-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 500\u0645\u0644", "conv": 24, "sp": 624.0, "rawBoxCost": 69.0417, "boxCost": 93.2063, "profit": 530.7937, "profitPct": 85.06, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-005-24", "qty": 1.0, "unitCost": 8.17256392294221, "lineCost": 8.1726}, {"child": "MB-032", "qty": 0.00072, "unitCost": 190.363636363636, "lineCost": 0.1371}, {"child": "MB-035", "qty": 0.00144, "unitCost": 186.397467572576, "lineCost": 0.2684}, {"child": "MB-036", "qty": 0.00072, "unitCost": 188.125, "lineCost": 0.1355}, {"child": "MB-043", "qty": 0.00072, "unitCost": 170.799180327869, "lineCost": 0.123}, {"child": "MB-044", "qty": 0.00072, "unitCost": 188.547619047619, "lineCost": 0.1358}, {"child": "MB-046", "qty": 0.00144, "unitCost": 112.599364069952, "lineCost": 0.1621}, {"child": "MB-047", "qty": 0.00072, "unitCost": 114.0, "lineCost": 0.0821}, {"child": "MB-048", "qty": 0.00072, "unitCost": 110.0, "lineCost": 0.0792}, {"child": "MB-049", "qty": 0.00072, "unitCost": 110.0, "lineCost": 0.0792}, {"child": "MB-050", "qty": 0.00072, "unitCost": 114.0, "lineCost": 0.0821}, {"child": "MB-051", "qty": 0.00072, "unitCost": 111.527777777778, "lineCost": 0.0803}, {"child": "MB-055", "qty": 0.00072, "unitCost": 214.316753926702, "lineCost": 0.1543}, {"child": "PET-001", "qty": 1.44, "unitCost": 61.9588348579682, "lineCost": 89.2207}, {"child": "PPH-001", "qty": 0.336, "unitCost": 67.4507739938081, "lineCost": 22.6635}, {"child": "ST-AQU-005-1", "qty": 24.0, "unitCost": 0.687817258883249, "lineCost": 16.5076}]}, {"code": "AQU-007-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0648\u064a\u0641 750\u0645\u0644", "conv": 24, "sp": 960.0, "rawBoxCost": 225.5461, "boxCost": 304.4872, "profit": 655.5128, "profitPct": 68.28, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-007-24", "qty": 1.0, "unitCost": 16.85, "lineCost": 16.85}, {"child": "MB-037", "qty": 0.02736, "unitCost": 109.956331877729, "lineCost": 3.0084}, {"child": "MB-201", "qty": 0.00258, "unitCost": 323.888888888889, "lineCost": 0.8356}, {"child": "MB-203", "qty": 0.00174, "unitCost": 344.815950920245, "lineCost": 0.6}, {"child": "MB-207", "qty": 0.00252, "unitCost": 415.0, "lineCost": 1.0458}, {"child": "MB-208", "qty": 0.00258, "unitCost": 241.666666666667, "lineCost": 0.6235}, {"child": "PET-001", "qty": 2.04, "unitCost": 61.9588348579682, "lineCost": 126.396}, {"child": "PPH-001", "qty": 0.912, "unitCost": 67.4507739938081, "lineCost": 61.5151}, {"child": "RUB-001", "qty": 0.036, "unitCost": 182.875, "lineCost": 6.5835}, {"child": "ST-AQU-007-1", "qty": 24.0, "unitCost": 0.337007874015748, "lineCost": 8.0882}]}, {"code": "AQU-009-12", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0645\u064a\u062c\u0627 1250\u0645\u0644", "conv": 12, "sp": 600.0, "rawBoxCost": 140.6836, "boxCost": 189.9229, "profit": 410.0771, "profitPct": 68.35, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-009", "qty": 1.0, "unitCost": 10.06, "lineCost": 10.06}, {"child": "MB-037", "qty": 0.01368, "unitCost": 109.956331877729, "lineCost": 1.5042}, {"child": "MB-201", "qty": 0.00168, "unitCost": 323.888888888889, "lineCost": 0.5441}, {"child": "MB-203", "qty": 0.00168, "unitCost": 344.815950920245, "lineCost": 0.5793}, {"child": "MB-208", "qty": 0.00168, "unitCost": 241.666666666667, "lineCost": 0.406}, {"child": "PET-001", "qty": 1.44, "unitCost": 61.9588348579682, "lineCost": 89.2207}, {"child": "PPH-001", "qty": 0.456, "unitCost": 67.4507739938081, "lineCost": 30.7576}, {"child": "RUB-001", "qty": 0.018, "unitCost": 182.875, "lineCost": 3.2917}, {"child": "ST-AQU-009-1", "qty": 12.0, "unitCost": 0.36, "lineCost": 4.32}]}, {"code": "AQU-009-60", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0645\u064a\u062c\u0627 1250\u0645\u0644 \u062a\u0639\u0628\u0626\u0629 60", "conv": 60, "sp": 0, "rawBoxCost": 554.0307, "boxCost": 747.9414, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-009-60", "qty": 1.0, "unitCost": 36.22, "lineCost": 36.22}, {"child": "MB-042", "qty": 0.036, "unitCost": 136.42723880597, "lineCost": 4.9114}, {"child": "MB-208", "qty": 0.0252, "unitCost": 241.666666666667, "lineCost": 6.09}, {"child": "PET-001", "qty": 7.2, "unitCost": 61.9588348579682, "lineCost": 446.1036}, {"child": "PPH-001", "qty": 0.9, "unitCost": 67.4507739938081, "lineCost": 60.7057}]}, {"code": "AQU-010-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 1 \u0644\u062a\u0631  ", "conv": 24, "sp": 912.0, "rawBoxCost": 184.7605, "boxCost": 249.4267, "profit": 662.5733, "profitPct": 72.65, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-010-24", "qty": 1.0, "unitCost": 14.8722988505747, "lineCost": 14.8723}, {"child": "MB-046", "qty": 0.0018, "unitCost": 112.599364069952, "lineCost": 0.2027}, {"child": "MB-047", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-048", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-049", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-050", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-051", "qty": 0.0018, "unitCost": 111.527777777778, "lineCost": 0.2008}, {"child": "PET-001", "qty": 2.04, "unitCost": 61.9588348579682, "lineCost": 126.396}, {"child": "PPH-001", "qty": 0.36, "unitCost": 67.4507739938081, "lineCost": 24.2823}, {"child": "ST-AQU-010-01", "qty": 24.0, "unitCost": 0.75, "lineCost": 18.0}]}, {"code": "AQU-010-60", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 1\u0644\u062a\u0631 \u062a\u0639\u0628\u0626\u0629 60  ", "conv": 60, "sp": 0, "rawBoxCost": 417.7039, "boxCost": 563.9003, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-010-60", "qty": 1.0, "unitCost": 29.522665855143, "lineCost": 29.5227}, {"child": "MB-042", "qty": 0.012, "unitCost": 136.42723880597, "lineCost": 1.6371}, {"child": "MB-043", "qty": 0.012, "unitCost": 170.799180327869, "lineCost": 2.0496}, {"child": "MB-046", "qty": 0.009, "unitCost": 112.599364069952, "lineCost": 1.0134}, {"child": "MB-201", "qty": 0.0086, "unitCost": 323.888888888889, "lineCost": 2.7854}, {"child": "MB-203", "qty": 0.0116, "unitCost": 344.815950920245, "lineCost": 3.9999}, {"child": "PET-001", "qty": 5.1, "unitCost": 61.9588348579682, "lineCost": 315.9901}, {"child": "PPH-001", "qty": 0.9, "unitCost": 67.4507739938081, "lineCost": 60.7057}]}, {"code": "AQU-011-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 1.5\u0644\u062a\u0631  ", "conv": 24, "sp": 1080.0, "rawBoxCost": 236.9945, "boxCost": 319.9426, "profit": 760.0574, "profitPct": 70.38, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-011", "qty": 1.0, "unitCost": 17.8988295069627, "lineCost": 17.8988}, {"child": "MB-046", "qty": 0.0018, "unitCost": 112.599364069952, "lineCost": 0.2027}, {"child": "MB-047", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-048", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-049", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-050", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-051", "qty": 0.0018, "unitCost": 111.527777777778, "lineCost": 0.2008}, {"child": "MB-203", "qty": 0.00168, "unitCost": 344.815950920245, "lineCost": 0.5793}, {"child": "MB-207", "qty": 0.00168, "unitCost": 415.0, "lineCost": 0.6972}, {"child": "MB-208", "qty": 0.00168, "unitCost": 241.666666666667, "lineCost": 0.406}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.36, "unitCost": 67.4507739938081, "lineCost": 24.2823}, {"child": "ST-AQU-011-1", "qty": 24.0, "unitCost": 0.561650717703349, "lineCost": 13.4796}]}, {"code": "AQU-011-60", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0633\u0643\u0648\u064a\u0631 1.5\u0644\u062a\u0631 \u062a\u0639\u0628\u0626\u0629 60  ", "conv": 60, "sp": 0, "rawBoxCost": 553.1466, "boxCost": 746.7479, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-011-60", "qty": 1.0, "unitCost": 35.3358778625954, "lineCost": 35.3359}, {"child": "MB-042", "qty": 0.036, "unitCost": 136.42723880597, "lineCost": 4.9114}, {"child": "MB-208", "qty": 0.0252, "unitCost": 241.666666666667, "lineCost": 6.09}, {"child": "PET-001", "qty": 7.2, "unitCost": 61.9588348579682, "lineCost": 446.1036}, {"child": "PPH-001", "qty": 0.9, "unitCost": 67.4507739938081, "lineCost": 60.7057}]}, {"code": "AQU-012-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0628\u064a\u0631\u0627\u0645\u064a\u062f\u0632 1\u0644\u062a\u0631  ", "conv": 24, "sp": 912.0, "rawBoxCost": 182.9166, "boxCost": 246.9374, "profit": 665.0626, "profitPct": 72.92, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-012-24", "qty": 1.0, "unitCost": 16.6283657199801, "lineCost": 16.6284}, {"child": "MB-046", "qty": 0.0018, "unitCost": 112.599364069952, "lineCost": 0.2027}, {"child": "MB-047", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-048", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-049", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-050", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-051", "qty": 0.0018, "unitCost": 111.527777777778, "lineCost": 0.2008}, {"child": "PET-001", "qty": 2.04, "unitCost": 61.9588348579682, "lineCost": 126.396}, {"child": "PPH-001", "qty": 0.36, "unitCost": 67.4507739938081, "lineCost": 24.2823}, {"child": "ST-AQU-012-01", "qty": 24.0, "unitCost": 0.6, "lineCost": 14.4}]}, {"code": "AQU-013-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0628\u064a\u0631\u0627\u0645\u064a\u062f\u0632 1.5\u0644\u062a\u0631  ", "conv": 24, "sp": 1080.0, "rawBoxCost": 244.298, "boxCost": 329.8023, "profit": 750.1977, "profitPct": 69.46, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-013-24", "qty": 1.0, "unitCost": 21.1644086886565, "lineCost": 21.1644}, {"child": "MB-046", "qty": 0.0018, "unitCost": 112.599364069952, "lineCost": 0.2027}, {"child": "MB-047", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-048", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-049", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-050", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-051", "qty": 0.0018, "unitCost": 111.527777777778, "lineCost": 0.2008}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.36, "unitCost": 67.4507739938081, "lineCost": 24.2823}, {"child": "ST-AQU-013-01", "qty": 24.0, "unitCost": 0.8, "lineCost": 19.2}]}, {"code": "AQU-014-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u062a\u0627\u0646\u0643 2.2 \u0644\u062a\u0631", "conv": 24, "sp": 1152.0, "rawBoxCost": 245.2786, "boxCost": 331.1261, "profit": 820.8739, "profitPct": 71.26, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-014-24", "qty": 1.0, "unitCost": 21.2565320197044, "lineCost": 21.2565}, {"child": "MB-046", "qty": 0.0018, "unitCost": 112.599364069952, "lineCost": 0.2027}, {"child": "MB-047", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-048", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-049", "qty": 0.0018, "unitCost": 110.0, "lineCost": 0.198}, {"child": "MB-050", "qty": 0.0018, "unitCost": 114.0, "lineCost": 0.2052}, {"child": "MB-051", "qty": 0.0018, "unitCost": 111.527777777778, "lineCost": 0.2008}, {"child": "MB-203", "qty": 0.00168, "unitCost": 344.815950920245, "lineCost": 0.5793}, {"child": "MB-207", "qty": 0.00168, "unitCost": 415.0, "lineCost": 0.6972}, {"child": "MB-208", "qty": 0.00336, "unitCost": 241.666666666667, "lineCost": 0.812}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.36, "unitCost": 67.4507739938081, "lineCost": 24.2823}, {"child": "ST-AQU-014-01", "qty": 24.0, "unitCost": 0.75, "lineCost": 18.0}]}, {"code": "AQU-015-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0627\u064a\u0646\u0631\u062c\u064a550 \u0645\u0644 ", "conv": 24, "sp": 840.0, "rawBoxCost": 235.8028, "boxCost": 318.3338, "profit": 521.6662, "profitPct": 62.1, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-015-24", "qty": 1.0, "unitCost": 10.25, "lineCost": 10.25}, {"child": "MB-046", "qty": 0.00392, "unitCost": 112.599364069952, "lineCost": 0.4414}, {"child": "MB-048", "qty": 0.00104, "unitCost": 110.0, "lineCost": 0.1144}, {"child": "MB-049", "qty": 0.00288, "unitCost": 110.0, "lineCost": 0.3168}, {"child": "MB-055", "qty": 0.00392, "unitCost": 214.316753926702, "lineCost": 0.8401}, {"child": "MB-202", "qty": 0.0032, "unitCost": 353.333333333333, "lineCost": 1.1307}, {"child": "MB-203", "qty": 0.0032, "unitCost": 344.815950920245, "lineCost": 1.1034}, {"child": "MB-208", "qty": 0.0032, "unitCost": 241.666666666667, "lineCost": 0.7733}, {"child": "PET-001", "qty": 1.2, "unitCost": 61.9588348579682, "lineCost": 74.3506}, {"child": "PPH-001", "qty": 0.3024, "unitCost": 67.4507739938081, "lineCost": 20.3971}, {"child": "PPR-001", "qty": 0.42, "unitCost": 80.1845484377229, "lineCost": 33.6775}, {"child": "RUB-001", "qty": 0.42, "unitCost": 182.875, "lineCost": 76.8075}, {"child": "ST-AQU-015-01", "qty": 24.0, "unitCost": 0.65, "lineCost": 15.6}]}, {"code": "AQU-016-24", "desc": "\u0632\u062c\u0627\u062c\u0629 \u0645\u064a\u0627\u0647 \u0627\u0643\u0648\u0627 \u0627\u064a\u0646\u0631\u062c\u064a750\u0645\u0644 ", "conv": 24, "sp": 750.0, "rawBoxCost": 251.0027, "boxCost": 338.8536, "profit": 411.1464, "profitPct": 54.82, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-AQU-016-24", "qty": 1.0, "unitCost": 11.35, "lineCost": 11.35}, {"child": "MB-046", "qty": 0.00294, "unitCost": 112.599364069952, "lineCost": 0.331}, {"child": "MB-048", "qty": 0.00078, "unitCost": 110.0, "lineCost": 0.0858}, {"child": "MB-049", "qty": 0.00216, "unitCost": 110.0, "lineCost": 0.2376}, {"child": "MB-050", "qty": 0.00282, "unitCost": 114.0, "lineCost": 0.3215}, {"child": "MB-055", "qty": 0.00294, "unitCost": 214.316753926702, "lineCost": 0.6301}, {"child": "MB-201", "qty": 0.0018, "unitCost": 323.888888888889, "lineCost": 0.583}, {"child": "MB-202", "qty": 0.0018, "unitCost": 353.333333333333, "lineCost": 0.636}, {"child": "MB-203", "qty": 0.0018, "unitCost": 344.815950920245, "lineCost": 0.6207}, {"child": "MB-207", "qty": 0.0018, "unitCost": 415.0, "lineCost": 0.747}, {"child": "PET-001", "qty": 1.44, "unitCost": 61.9588348579682, "lineCost": 89.2207}, {"child": "PPH-001", "qty": 0.2988, "unitCost": 67.4507739938081, "lineCost": 20.1543}, {"child": "PPR-001", "qty": 0.42, "unitCost": 80.1845484377229, "lineCost": 33.6775}, {"child": "RUB-001", "qty": 0.42, "unitCost": 182.875, "lineCost": 76.8075}, {"child": "ST-AQU-016-01", "qty": 24.0, "unitCost": 0.65, "lineCost": 15.6}]}, {"code": "BRB-001-06", "desc": "\u062e\u0628\u0627\u0632\u0629 \u0628\u0631\u064a\u062f\u0644\u064a", "conv": 6, "sp": 1164.0, "rawBoxCost": 454.4195, "boxCost": 613.4663, "profit": 550.5337, "profitPct": 47.3, "hasBOM": true, "missing": [], "bomLines": [{"child": "CI-BRB-001-01", "qty": 6.0, "unitCost": 15.0, "lineCost": 90.0}, {"child": "MB-018", "qty": 0.01785, "unitCost": 100.427215189873, "lineCost": 1.7926}, {"child": "MB-030", "qty": 0.01785, "unitCost": 223.736576152305, "lineCost": 3.9937}, {"child": "MB-040", "qty": 0.01785, "unitCost": 114.00269541779, "lineCost": 2.0349}, {"child": "MB-046", "qty": 0.01785, "unitCost": 112.599364069952, "lineCost": 2.0099}, {"child": "MB-047", "qty": 0.01785, "unitCost": 114.0, "lineCost": 2.0349}, {"child": "MB-104", "qty": 0.01785, "unitCost": 265.0, "lineCost": 4.7303}, {"child": "PPK-001", "qty": 3.57, "unitCost": 75.3691770186335, "lineCost": 269.068}, {"child": "PPR-001", "qty": 0.81, "unitCost": 80.1845484377229, "lineCost": 64.9495}, {"child": "PSH-BRB-001-06", "qty": 0.2, "unitCost": 69.0286259541985, "lineCost": 13.8057}]}, {"code": "BUB-001-12", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0646\u0627\u062f\u064a\u0644 \u0628\u0627\u0628\u0644\u0632 ", "conv": 12, "sp": 1470.0, "rawBoxCost": 506.4617, "boxCost": 683.7233, "profit": 786.2767, "profitPct": 53.49, "hasBOM": true, "missing": [], "bomLines": [{"child": "ACC-BUB-001", "qty": 12.0, "unitCost": 0.57741935483871, "lineCost": 6.929}, {"child": "CE-BUB-001-12", "qty": 1.0, "unitCost": 16.65395, "lineCost": 16.6539}, {"child": "CI-BUB-001-01", "qty": 12.0, "unitCost": 5.79893488105109, "lineCost": 69.5872}, {"child": "MB-037", "qty": 0.014, "unitCost": 109.956331877729, "lineCost": 1.5394}, {"child": "MB-039", "qty": 0.014, "unitCost": 96.0381355932203, "lineCost": 1.3445}, {"child": "PPS-001", "qty": 4.74, "unitCost": 86.5839025839026, "lineCost": 410.4077}]}, {"code": "BUB-001-60", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0646\u0627\u062f\u064a\u0644 \u0628\u0627\u0628\u0644\u0632 \u062a\u0639\u0628\u0626\u0629 60", "conv": 60, "sp": 0, "rawBoxCost": 2189.0117, "boxCost": 2955.1658, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "ACC-BUB-001", "qty": 60.0, "unitCost": 0.57741935483871, "lineCost": 34.6452}, {"child": "CE-BUB-001-60", "qty": 1.0, "unitCost": 82.16, "lineCost": 82.16}, {"child": "MB-039", "qty": 0.21, "unitCost": 96.0381355932203, "lineCost": 20.168}, {"child": "PPS-001", "qty": 23.7, "unitCost": 86.5839025839026, "lineCost": 2052.0385}]}, {"code": "CYL-001-60", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 \u0641\u0631\u062f\u0627\u0646\u064a 300 \u0645\u0644 \u0623\u0644\u0648\u0627\u0646", "conv": 60, "sp": 1350.0, "rawBoxCost": 318.3833, "boxCost": 429.8175, "profit": 920.1825, "profitPct": 68.16, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-001-60", "qty": 1.0, "unitCost": 18.11, "lineCost": 18.11}, {"child": "MB-024", "qty": 0.0066, "unitCost": 112.5, "lineCost": 0.7425}, {"child": "MB-030", "qty": 0.0066, "unitCost": 223.736576152305, "lineCost": 1.4767}, {"child": "MB-037", "qty": 0.0066, "unitCost": 109.956331877729, "lineCost": 0.7257}, {"child": "MB-038", "qty": 0.0066, "unitCost": 185.610749185668, "lineCost": 1.225}, {"child": "MB-047", "qty": 0.0066, "unitCost": 114.0, "lineCost": 0.7524}, {"child": "MB-052", "qty": 0.0066, "unitCost": 100.0, "lineCost": 0.66}, {"child": "PPS-001", "qty": 2.1, "unitCost": 86.5839025839026, "lineCost": 181.8262}, {"child": "PPS-101", "qty": 1.32, "unitCost": 71.5607979734009, "lineCost": 94.4603}, {"child": "ST-CYL-001", "qty": 60.0, "unitCost": 0.30674130251595, "lineCost": 18.4045}]}, {"code": "CYL-002-60", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 \u0641\u0631\u062f\u0627\u0646\u064a 450 \u0645\u0644 \u0623\u0644\u0648\u0627\u0646", "conv": 60, "sp": 1500.0, "rawBoxCost": 387.9487, "boxCost": 523.7307, "profit": 976.2693, "profitPct": 65.08, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-002-60", "qty": 1.0, "unitCost": 20.83, "lineCost": 20.83}, {"child": "MB-024", "qty": 0.0066, "unitCost": 112.5, "lineCost": 0.7425}, {"child": "MB-030", "qty": 0.0066, "unitCost": 223.736576152305, "lineCost": 1.4767}, {"child": "MB-037", "qty": 0.0066, "unitCost": 109.956331877729, "lineCost": 0.7257}, {"child": "MB-038", "qty": 0.0066, "unitCost": 185.610749185668, "lineCost": 1.225}, {"child": "MB-047", "qty": 0.0066, "unitCost": 114.0, "lineCost": 0.7524}, {"child": "MB-052", "qty": 0.0066, "unitCost": 100.0, "lineCost": 0.66}, {"child": "PPS-001", "qty": 2.82, "unitCost": 86.5839025839026, "lineCost": 244.1666}, {"child": "PPS-101", "qty": 1.32, "unitCost": 71.5607979734009, "lineCost": 94.4603}, {"child": "ST-CYL-002", "qty": 60.0, "unitCost": 0.381825506214412, "lineCost": 22.9095}]}, {"code": "CYL-003-60", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 \u0641\u0631\u062f\u0627\u0646\u064a 750 \u0645\u0644 \u0627\u0644\u0648\u0627\u0646", "conv": 60, "sp": 2250.0, "rawBoxCost": 548.5552, "boxCost": 740.5495, "profit": 1509.4505, "profitPct": 67.09, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-003-60", "qty": 1.0, "unitCost": 26.69, "lineCost": 26.69}, {"child": "MB-024", "qty": 0.0066, "unitCost": 112.5, "lineCost": 0.7425}, {"child": "MB-030", "qty": 0.0066, "unitCost": 223.736576152305, "lineCost": 1.4767}, {"child": "MB-037", "qty": 0.0066, "unitCost": 109.956331877729, "lineCost": 0.7257}, {"child": "MB-038", "qty": 0.0066, "unitCost": 185.610749185668, "lineCost": 1.225}, {"child": "MB-047", "qty": 0.0066, "unitCost": 114.0, "lineCost": 0.7524}, {"child": "MB-052", "qty": 0.0066, "unitCost": 100.0, "lineCost": 0.66}, {"child": "PPS-001", "qty": 4.56, "unitCost": 86.5839025839026, "lineCost": 394.8226}, {"child": "PPS-101", "qty": 1.32, "unitCost": 71.5607979734009, "lineCost": 94.4603}, {"child": "ST-CYL-003", "qty": 60.0, "unitCost": 0.45, "lineCost": 27.0}]}, {"code": "CYL-011-60", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 \u0641\u0631\u062f\u0627\u0646\u064a 300 \u0645\u0644 \u062e\u0634\u0628\u064a", "conv": 60, "sp": 0, "rawBoxCost": 453.8778, "boxCost": 612.735, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-001-60", "qty": 1.0, "unitCost": 18.11, "lineCost": 18.11}, {"child": "MB-005", "qty": 0.0396, "unitCost": 123.472527472527, "lineCost": 4.8895}, {"child": "PAS-001", "qty": 1.32, "unitCost": 50.1727167246809, "lineCost": 66.228}, {"child": "PPS-001", "qty": 2.1, "unitCost": 86.5839025839026, "lineCost": 181.8262}, {"child": "ST-CYL-001", "qty": 60.0, "unitCost": 0.30674130251595, "lineCost": 18.4045}, {"child": "ST-WOD-001", "qty": 60.0, "unitCost": 2.74032731143599, "lineCost": 164.4196}]}, {"code": "CYL-012-60", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 \u0641\u0631\u062f\u0627\u0646\u064a 450 \u0645\u0644 \u062e\u0634\u0628\u064a", "conv": 60, "sp": 0, "rawBoxCost": 523.4432, "boxCost": 706.6483, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-002-60", "qty": 1.0, "unitCost": 20.83, "lineCost": 20.83}, {"child": "MB-005", "qty": 0.0396, "unitCost": 123.472527472527, "lineCost": 4.8895}, {"child": "PAS-001", "qty": 1.32, "unitCost": 50.1727167246809, "lineCost": 66.228}, {"child": "PPS-001", "qty": 2.82, "unitCost": 86.5839025839026, "lineCost": 244.1666}, {"child": "ST-CYL-002", "qty": 60.0, "unitCost": 0.381825506214412, "lineCost": 22.9095}, {"child": "ST-WOD-001", "qty": 60.0, "unitCost": 2.74032731143599, "lineCost": 164.4196}]}, {"code": "CYL-101-12", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 9 \u0641\u064a 1 \u0623\u0644\u0648\u0627\u0646", "conv": 12, "sp": 0, "rawBoxCost": 835.5892, "boxCost": 1128.0454, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-101-12", "qty": 1.0, "unitCost": 31.476370510397, "lineCost": 31.4764}, {"child": "CI-CYL-101-01", "qty": 12.0, "unitCost": 9.92331288343558, "lineCost": 119.0798}, {"child": "MB-024", "qty": 0.01188, "unitCost": 112.5, "lineCost": 1.3365}, {"child": "MB-030", "qty": 0.01188, "unitCost": 223.736576152305, "lineCost": 2.658}, {"child": "MB-037", "qty": 0.01188, "unitCost": 109.956331877729, "lineCost": 1.3063}, {"child": "MB-038", "qty": 0.01188, "unitCost": 185.610749185668, "lineCost": 2.2051}, {"child": "MB-047", "qty": 0.01188, "unitCost": 114.0, "lineCost": 1.3543}, {"child": "MB-052", "qty": 0.01188, "unitCost": 100.0, "lineCost": 1.188}, {"child": "PPS-001", "qty": 5.688, "unitCost": 86.5839025839026, "lineCost": 492.4892}, {"child": "PPS-101", "qty": 2.376, "unitCost": 71.5607979734009, "lineCost": 170.0285}, {"child": "ST-CYL-101", "qty": 12.0, "unitCost": 1.03892312682452, "lineCost": 12.4671}]}, {"code": "CYL-102-12", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 3 \u0641\u064a 1 \u0623\u0644\u0648\u0627\u0646", "conv": 12, "sp": 0, "rawBoxCost": 287.5897, "boxCost": 388.2461, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-102-12", "qty": 1.0, "unitCost": 14.2011100917431, "lineCost": 14.2011}, {"child": "CI-CYL-102-01", "qty": 12.0, "unitCost": 4.1, "lineCost": 49.2}, {"child": "MB-024", "qty": 0.00396, "unitCost": 112.5, "lineCost": 0.4455}, {"child": "MB-030", "qty": 0.00396, "unitCost": 223.736576152305, "lineCost": 0.886}, {"child": "MB-037", "qty": 0.00396, "unitCost": 109.956331877729, "lineCost": 0.4354}, {"child": "MB-038", "qty": 0.00396, "unitCost": 185.610749185668, "lineCost": 0.735}, {"child": "MB-047", "qty": 0.00396, "unitCost": 114.0, "lineCost": 0.4514}, {"child": "MB-052", "qty": 0.00396, "unitCost": 100.0, "lineCost": 0.396}, {"child": "PPS-001", "qty": 1.896, "unitCost": 86.5839025839026, "lineCost": 164.1631}, {"child": "PPS-101", "qty": 0.792, "unitCost": 71.5607979734009, "lineCost": 56.6762}]}, {"code": "CYL-103-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0631\u0628\u0627\u0639\u064a \u0633\u064a\u0644\u064a\u0648 300 \u0645\u0644 \u0627\u0644\u0648\u0627\u0646", "conv": 12, "sp": 1350.0, "rawBoxCost": 407.4727, "boxCost": 550.0881, "profit": 799.9119, "profitPct": 59.25, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-103-12", "qty": 1.0, "unitCost": 22.6615864022663, "lineCost": 22.6616}, {"child": "MB-030", "qty": 0.0066, "unitCost": 223.736576152305, "lineCost": 1.4767}, {"child": "MB-037", "qty": 0.0066, "unitCost": 109.956331877729, "lineCost": 0.7257}, {"child": "MB-038", "qty": 0.0066, "unitCost": 185.610749185668, "lineCost": 1.225}, {"child": "MB-040", "qty": 0.0066, "unitCost": 114.00269541779, "lineCost": 0.7524}, {"child": "MB-046", "qty": 0.0066, "unitCost": 112.599364069952, "lineCost": 0.7432}, {"child": "MB-047", "qty": 0.0066, "unitCost": 114.0, "lineCost": 0.7524}, {"child": "PPS-001", "qty": 3.336, "unitCost": 86.5839025839026, "lineCost": 288.8439}, {"child": "PPS-101", "qty": 1.056, "unitCost": 71.5607979734009, "lineCost": 75.5682}, {"child": "ST-CYL-001", "qty": 48.0, "unitCost": 0.30674130251595, "lineCost": 14.7236}]}, {"code": "CYL-104-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0631\u0628\u0627\u0639\u064a \u0633\u064a\u0644\u064a\u0648 450 \u0645\u0644 \u0627\u0644\u0648\u0627\u0646", "conv": 12, "sp": 1575.0, "rawBoxCost": 464.7535, "boxCost": 627.4172, "profit": 947.5828, "profitPct": 60.16, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-104-12", "qty": 1.0, "unitCost": 26.4660669077758, "lineCost": 26.4661}, {"child": "MB-030", "qty": 0.0066, "unitCost": 223.736576152305, "lineCost": 1.4767}, {"child": "MB-037", "qty": 0.0066, "unitCost": 109.956331877729, "lineCost": 0.7257}, {"child": "MB-038", "qty": 0.0066, "unitCost": 185.610749185668, "lineCost": 1.225}, {"child": "MB-040", "qty": 0.0066, "unitCost": 114.00269541779, "lineCost": 0.7524}, {"child": "MB-046", "qty": 0.0066, "unitCost": 112.599364069952, "lineCost": 0.7432}, {"child": "MB-047", "qty": 0.0066, "unitCost": 114.0, "lineCost": 0.7524}, {"child": "PPS-001", "qty": 3.912, "unitCost": 86.5839025839026, "lineCost": 338.7162}, {"child": "PPS-101", "qty": 1.056, "unitCost": 71.5607979734009, "lineCost": 75.5682}, {"child": "ST-CYL-002", "qty": 48.0, "unitCost": 0.381825506214412, "lineCost": 18.3276}]}, {"code": "CYL-105-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0628\u0647\u0627\u0631\u0627\u062a  \u0633\u064a\u0644\u064a\u0648 2 \u062f\u0648\u0631 \u0623\u0644\u0648\u0627\u0646", "conv": 12, "sp": 3000.0, "rawBoxCost": 816.481, "boxCost": 1102.2494, "profit": 1897.7506, "profitPct": 63.26, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-105-12", "qty": 1.0, "unitCost": 24.5743792063124, "lineCost": 24.5744}, {"child": "MB-030", "qty": 0.03324, "unitCost": 223.736576152305, "lineCost": 7.437}, {"child": "MB-037", "qty": 0.03324, "unitCost": 109.956331877729, "lineCost": 3.6549}, {"child": "MB-038", "qty": 0.03324, "unitCost": 185.610749185668, "lineCost": 6.1697}, {"child": "MB-040", "qty": 0.03324, "unitCost": 114.00269541779, "lineCost": 3.7894}, {"child": "MB-046", "qty": 0.03324, "unitCost": 112.599364069952, "lineCost": 3.7428}, {"child": "MB-047", "qty": 0.03324, "unitCost": 114.0, "lineCost": 3.7894}, {"child": "PAS-001", "qty": 4.8, "unitCost": 50.1727167246809, "lineCost": 240.829}, {"child": "PPS-001", "qty": 3.864, "unitCost": 86.5839025839026, "lineCost": 334.5602}, {"child": "PPS-101", "qty": 1.848, "unitCost": 71.5607979734009, "lineCost": 132.2444}, {"child": "ST-CYL-001", "qty": 36.0, "unitCost": 0.30674130251595, "lineCost": 11.0427}, {"child": "ST-CYL-002", "qty": 36.0, "unitCost": 0.381825506214412, "lineCost": 13.7457}, {"child": "ST-CYL-003", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}, {"child": "ST-CYL-101", "qty": 12.0, "unitCost": 1.03892312682452, "lineCost": 12.4671}, {"child": "ST-CYL-105", "qty": 12.0, "unitCost": 1.08619119878604, "lineCost": 13.0343}]}, {"code": "CYL-201-12", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 9 \u0641\u064a 1 \u062e\u0634\u0628\u064a", "conv": 12, "sp": 3375.0, "rawBoxCost": 1079.4793, "boxCost": 1457.2971, "profit": 1917.7029, "profitPct": 56.82, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-101-12", "qty": 1.0, "unitCost": 31.476370510397, "lineCost": 31.4764}, {"child": "CI-CYL-101-01", "qty": 12.0, "unitCost": 9.92331288343558, "lineCost": 119.0798}, {"child": "MB-005", "qty": 0.07128, "unitCost": 123.472527472527, "lineCost": 8.8011}, {"child": "PAS-001", "qty": 2.376, "unitCost": 50.1727167246809, "lineCost": 119.2104}, {"child": "PPS-001", "qty": 5.688, "unitCost": 86.5839025839026, "lineCost": 492.4892}, {"child": "ST-CYL-101", "qty": 12.0, "unitCost": 1.03892312682452, "lineCost": 12.4671}, {"child": "ST-WOD-001", "qty": 108.0, "unitCost": 2.74032731143599, "lineCost": 295.9553}]}, {"code": "CYL-202-12", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0647\u0627\u0631\u0627\u062a \u0633\u064a\u0644\u064a\u0648 3 \u0641\u064a 1 \u062e\u0634\u0628\u064a", "conv": 12, "sp": 1200.0, "rawBoxCost": 368.8865, "boxCost": 497.9968, "profit": 702.0032, "profitPct": 58.5, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-102-12", "qty": 1.0, "unitCost": 14.2011100917431, "lineCost": 14.2011}, {"child": "CI-CYL-102-01", "qty": 12.0, "unitCost": 4.1, "lineCost": 49.2}, {"child": "MB-005", "qty": 0.02376, "unitCost": 123.472527472527, "lineCost": 2.9337}, {"child": "PAS-001", "qty": 0.792, "unitCost": 50.1727167246809, "lineCost": 39.7368}, {"child": "PPS-001", "qty": 1.896, "unitCost": 86.5839025839026, "lineCost": 164.1631}, {"child": "ST-WOD-001", "qty": 36.0, "unitCost": 2.74032731143599, "lineCost": 98.6518}]}, {"code": "CYL-203-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0631\u0628\u0627\u0639\u064a \u0633\u064a\u0644\u064a\u0648 300 \u0645\u0644 \u062e\u0634\u0628\u064a", "conv": 12, "sp": 0, "rawBoxCost": 455.3398, "boxCost": 614.7087, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-103-12", "qty": 1.0, "unitCost": 22.6615864022663, "lineCost": 22.6616}, {"child": "MB-005", "qty": 0.0396, "unitCost": 123.472527472527, "lineCost": 4.8895}, {"child": "PAS-001", "qty": 2.712, "unitCost": 50.1727167246809, "lineCost": 136.0684}, {"child": "PPS-001", "qty": 1.68, "unitCost": 86.5839025839026, "lineCost": 145.461}, {"child": "ST-CYL-001", "qty": 48.0, "unitCost": 0.30674130251595, "lineCost": 14.7236}, {"child": "ST-WOD-001", "qty": 48.0, "unitCost": 2.74032731143599, "lineCost": 131.5357}]}, {"code": "CYL-204-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0631\u0628\u0627\u0639\u064a \u0633\u064a\u0644\u064a\u0648 450 \u0645\u0644 \u062e\u0634\u0628\u064a", "conv": 12, "sp": 0, "rawBoxCost": 512.6206, "boxCost": 692.0378, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-104-12", "qty": 1.0, "unitCost": 26.4660669077758, "lineCost": 26.4661}, {"child": "MB-005", "qty": 0.0396, "unitCost": 123.472527472527, "lineCost": 4.8895}, {"child": "PAS-001", "qty": 2.712, "unitCost": 50.1727167246809, "lineCost": 136.0684}, {"child": "PPS-001", "qty": 2.256, "unitCost": 86.5839025839026, "lineCost": 195.3333}, {"child": "ST-CYL-002", "qty": 48.0, "unitCost": 0.381825506214412, "lineCost": 18.3276}, {"child": "ST-WOD-001", "qty": 48.0, "unitCost": 2.74032731143599, "lineCost": 131.5357}]}, {"code": "CYL-205-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0628\u0647\u0627\u0631\u0627\u062a  \u0633\u064a\u0644\u064a\u0648 2 \u062f\u0648\u0631 \u062e\u0634\u0628\u064a", "conv": 12, "sp": 0, "rawBoxCost": 1003.1855, "boxCost": 1354.3004, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-105-12", "qty": 1.0, "unitCost": 24.5743792063124, "lineCost": 24.5744}, {"child": "MB-005", "qty": 0.19944, "unitCost": 123.472527472527, "lineCost": 24.6254}, {"child": "PAS-001", "qty": 6.648, "unitCost": 50.1727167246809, "lineCost": 333.5482}, {"child": "PPS-001", "qty": 3.864, "unitCost": 86.5839025839026, "lineCost": 334.5602}, {"child": "ST-CYL-001", "qty": 36.0, "unitCost": 0.30674130251595, "lineCost": 11.0427}, {"child": "ST-CYL-002", "qty": 36.0, "unitCost": 0.381825506214412, "lineCost": 13.7457}, {"child": "ST-CYL-003", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}, {"child": "ST-CYL-101", "qty": 12.0, "unitCost": 1.03892312682452, "lineCost": 12.4671}, {"child": "ST-CYL-105", "qty": 12.0, "unitCost": 1.08619119878604, "lineCost": 13.0343}, {"child": "ST-WOD-001", "qty": 84.0, "unitCost": 2.74032731143599, "lineCost": 230.1875}]}, {"code": "CYL-305-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0628\u0647\u0627\u0631\u0627\u062a  \u0633\u064a\u0644\u064a\u0648 2 \u062f\u0648\u0631 \u0627\u0628\u064a\u0636", "conv": 12, "sp": 0, "rawBoxCost": 853.2797, "boxCost": 1151.9276, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-CYL-105-12", "qty": 1.0, "unitCost": 24.5743792063124, "lineCost": 24.5744}, {"child": "MB-039", "qty": 0.7488, "unitCost": 96.0381355932203, "lineCost": 71.9134}, {"child": "PAS-001", "qty": 5.808, "unitCost": 50.1727167246809, "lineCost": 291.4031}, {"child": "PPS-001", "qty": 3.864, "unitCost": 86.5839025839026, "lineCost": 334.5602}, {"child": "PPS-101", "qty": 1.05, "unitCost": 71.5607979734009, "lineCost": 75.1388}, {"child": "ST-CYL-001", "qty": 36.0, "unitCost": 0.30674130251595, "lineCost": 11.0427}, {"child": "ST-CYL-002", "qty": 36.0, "unitCost": 0.381825506214412, "lineCost": 13.7457}, {"child": "ST-CYL-003", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}, {"child": "ST-CYL-101", "qty": 12.0, "unitCost": 1.03892312682452, "lineCost": 12.4671}, {"child": "ST-CYL-105", "qty": 12.0, "unitCost": 1.08619119878604, "lineCost": 13.0343}]}, {"code": "DIS-001-24", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u0645\u0641\u0631\u062f \u0635\u063a\u064a\u0631 800", "conv": 24, "sp": 1353.0, "rawBoxCost": 416.0988, "boxCost": 561.7334, "profit": 791.2666, "profitPct": 58.48, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.3528, "unitCost": 106.131313131313, "lineCost": 37.4431}, {"child": "CE-DIS-001-24", "qty": 1.0, "unitCost": 13.4849210822999, "lineCost": 13.4849}, {"child": "CI-DIS-001-01", "qty": 24.0, "unitCost": 2.13404932378679, "lineCost": 51.2172}, {"child": "MB-001", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-003", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-030", "qty": 0.0018, "unitCost": 223.736576152305, "lineCost": 0.4027}, {"child": "MB-056", "qty": 0.0018, "unitCost": 120.0, "lineCost": 0.216}, {"child": "PPS-001", "qty": 3.36, "unitCost": 86.5839025839026, "lineCost": 290.9219}, {"child": "RUB-001", "qty": 0.12, "unitCost": 182.875, "lineCost": 21.945}]}, {"code": "DIS-002-24", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u0645\u0641\u0631\u062f \u0648\u0633\u0637 1000", "conv": 24, "sp": 1585.0, "rawBoxCost": 469.242, "boxCost": 633.4767, "profit": 951.5233, "profitPct": 60.03, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.3528, "unitCost": 106.131313131313, "lineCost": 37.4431}, {"child": "CE-DIS-002-24", "qty": 1.0, "unitCost": 14.9388507183011, "lineCost": 14.9389}, {"child": "CI-DIS-002-01", "qty": 24.0, "unitCost": 1.95, "lineCost": 46.8}, {"child": "MB-001", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-003", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-030", "qty": 0.0018, "unitCost": 223.736576152305, "lineCost": 0.4027}, {"child": "MB-056", "qty": 0.0018, "unitCost": 120.0, "lineCost": 0.216}, {"child": "PPS-001", "qty": 4.008, "unitCost": 86.5839025839026, "lineCost": 347.0283}, {"child": "RUB-001", "qty": 0.12, "unitCost": 182.875, "lineCost": 21.945}]}, {"code": "DIS-003-24", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u0645\u0641\u0631\u062f \u0643\u0628\u064a\u0631 1300", "conv": 24, "sp": 1813.0, "rawBoxCost": 555.582, "boxCost": 750.0357, "profit": 1062.9643, "profitPct": 58.63, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.3528, "unitCost": 106.131313131313, "lineCost": 37.4431}, {"child": "CE-DIS-003-24", "qty": 1.0, "unitCost": 16.8807385229541, "lineCost": 16.8807}, {"child": "CI-DIS-003-01", "qty": 24.0, "unitCost": 2.52274045870728, "lineCost": 60.5458}, {"child": "MB-001", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-003", "qty": 0.0018, "unitCost": 130.0, "lineCost": 0.234}, {"child": "MB-030", "qty": 0.0018, "unitCost": 223.736576152305, "lineCost": 0.4027}, {"child": "MB-056", "qty": 0.0018, "unitCost": 120.0, "lineCost": 0.216}, {"child": "PPS-001", "qty": 4.824, "unitCost": 86.5839025839026, "lineCost": 417.6807}, {"child": "RUB-001", "qty": 0.12, "unitCost": 182.875, "lineCost": 21.945}]}, {"code": "DIS-101-12", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u062b\u0644\u0627\u062b\u064a  \u0635\u063a\u064a\u0631800", "conv": 12, "sp": 1981.0, "rawBoxCost": 590.5532, "boxCost": 797.2468, "profit": 1183.7532, "profitPct": 59.76, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.5292, "unitCost": 106.131313131313, "lineCost": 56.1647}, {"child": "CE-DIS-101-12", "qty": 1.0, "unitCost": 11.8580031695721, "lineCost": 11.858}, {"child": "CI-DIS-101-03", "qty": 12.0, "unitCost": 4.3, "lineCost": 51.6}, {"child": "MB-001", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-003", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-030", "qty": 0.0027, "unitCost": 223.736576152305, "lineCost": 0.6041}, {"child": "MB-056", "qty": 0.0027, "unitCost": 120.0, "lineCost": 0.324}, {"child": "PPS-001", "qty": 5.04, "unitCost": 86.5839025839026, "lineCost": 436.3829}, {"child": "RUB-001", "qty": 0.18, "unitCost": 182.875, "lineCost": 32.9175}]}, {"code": "DIS-102-12", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u062b\u0644\u0627\u062b\u064a \u0648\u0633\u0637 1000", "conv": 12, "sp": 2313.0, "rawBoxCost": 670.0087, "boxCost": 904.5117, "profit": 1408.4883, "profitPct": 60.89, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.5292, "unitCost": 106.131313131313, "lineCost": 56.1647}, {"child": "CE-DIS-102-12", "qty": 1.0, "unitCost": 13.149255774166, "lineCost": 13.1493}, {"child": "CI-DIS-102-03", "qty": 12.0, "unitCost": 3.80039363621453, "lineCost": 45.6047}, {"child": "MB-001", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-003", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-030", "qty": 0.0027, "unitCost": 223.736576152305, "lineCost": 0.6041}, {"child": "MB-056", "qty": 0.0027, "unitCost": 120.0, "lineCost": 0.324}, {"child": "PPS-001", "qty": 6.012, "unitCost": 86.5839025839026, "lineCost": 520.5424}, {"child": "RUB-001", "qty": 0.18, "unitCost": 182.875, "lineCost": 32.9175}]}, {"code": "DIS-103-12", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u062b\u0644\u0627\u062b\u064a \u0643\u0628\u064a\u0631 1300", "conv": 12, "sp": 2638.0, "rawBoxCost": 783.8231, "boxCost": 1058.1612, "profit": 1579.8388, "profitPct": 59.89, "hasBOM": true, "missing": [], "bomLines": [{"child": "MB-030", "qty": 0.0027, "unitCost": 223.736576152305, "lineCost": 0.6041}, {"child": "MB-056", "qty": 0.0027, "unitCost": 120.0, "lineCost": 0.324}, {"child": "PPS-001", "qty": 7.236, "unitCost": 86.5839025839026, "lineCost": 626.5211}, {"child": "RUB-001", "qty": 0.18, "unitCost": 182.875, "lineCost": 32.9175}, {"child": "ABS-001", "qty": 0.5292, "unitCost": 106.131313131313, "lineCost": 56.1647}, {"child": "CE-DIS-103-12", "qty": 1.0, "unitCost": 14.4853191489362, "lineCost": 14.4853}, {"child": "CI-DIS-103-03", "qty": 12.0, "unitCost": 4.342034847542, "lineCost": 52.1044}, {"child": "MB-001", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-003", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}]}, {"code": "DIS-104-12", "desc": "\u0639\u0644\u0628\u0629 \u062f\u064a\u0633\u0643\u0641\u0631\u064a \u062b\u0644\u0627\u062b\u064a \u0645\u0634\u0643\u0644", "conv": 12, "sp": 2250.0, "rawBoxCost": 709.2605, "boxCost": 957.5017, "profit": 1292.4983, "profitPct": 57.44, "hasBOM": true, "missing": [], "bomLines": [{"child": "ABS-001", "qty": 0.5292, "unitCost": 106.131313131313, "lineCost": 56.1647}, {"child": "CE-DIS-104-12", "qty": 1.0, "unitCost": 20.1639857045609, "lineCost": 20.164}, {"child": "CI-DIS-104-03", "qty": 12.0, "unitCost": 5.88072850424443, "lineCost": 70.5687}, {"child": "MB-001", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-003", "qty": 0.0027, "unitCost": 130.0, "lineCost": 0.351}, {"child": "MB-030", "qty": 0.0027, "unitCost": 223.736576152305, "lineCost": 0.6041}, {"child": "MB-056", "qty": 0.0027, "unitCost": 120.0, "lineCost": 0.324}, {"child": "PPS-001", "qty": 6.096, "unitCost": 86.5839025839026, "lineCost": 527.8155}, {"child": "RUB-001", "qty": 0.18, "unitCost": 182.875, "lineCost": 32.9175}]}, {"code": "EVO-001-12", "desc": "\u0639\u0644\u0628\u0629 \u062a\u062e\u0632\u064a\u0646 \u0627\u064a\u0641\u0648 \u0635\u063a\u064a\u0631\u0629 1.25 \u0644\u062a\u0631", "conv": 12, "sp": 494.0, "rawBoxCost": 152.4057, "boxCost": 205.7477, "profit": 288.2523, "profitPct": 58.35, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-EVO-001-12", "qty": 1.0, "unitCost": 11.1740778170793, "lineCost": 11.1741}, {"child": "MB-039", "qty": 0.00432, "unitCost": 96.0381355932203, "lineCost": 0.4149}, {"child": "MB-040", "qty": 0.00288, "unitCost": 114.00269541779, "lineCost": 0.3283}, {"child": "MB-101", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-102", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-103", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-104", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-105", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "PPH-001", "qty": 0.70992, "unitCost": 67.4507739938081, "lineCost": 47.8847}, {"child": "PPR-001", "qty": 1.08, "unitCost": 80.1845484377229, "lineCost": 86.5993}, {"child": "ST-EVO-001", "qty": 12.0, "unitCost": 0.341363636363636, "lineCost": 4.0964}]}, {"code": "EVO-002-12", "desc": "\u0639\u0644\u0628\u0629 \u062a\u062e\u0632\u064a\u0646 \u0627\u064a\u0641\u0648 \u0648\u0633\u0637 2 \u0644\u062a\u0631", "conv": 12, "sp": 569.0, "rawBoxCost": 185.5242, "boxCost": 250.4577, "profit": 318.5423, "profitPct": 55.98, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-EVO-002-12", "qty": 1.0, "unitCost": 14.7225758143018, "lineCost": 14.7226}, {"child": "MB-039", "qty": 0.00432, "unitCost": 96.0381355932203, "lineCost": 0.4149}, {"child": "MB-040", "qty": 0.00288, "unitCost": 114.00269541779, "lineCost": 0.3283}, {"child": "MB-101", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-102", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-103", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-104", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-105", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "PPH-001", "qty": 0.70992, "unitCost": 67.4507739938081, "lineCost": 47.8847}, {"child": "PPR-001", "qty": 1.44, "unitCost": 80.1845484377229, "lineCost": 115.4657}, {"child": "ST-EVO-002", "qty": 12.0, "unitCost": 0.4, "lineCost": 4.8}]}, {"code": "EVO-003-12", "desc": "\u0639\u0644\u0628\u0629 \u062a\u062e\u0632\u064a\u0646 \u0627\u064a\u0641\u0648 \u0643\u0628\u064a\u0631 2.75 \u0644\u062a\u0631", "conv": 12, "sp": 663.0, "rawBoxCost": 227.0804, "boxCost": 306.5585, "profit": 356.4415, "profitPct": 53.76, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-EVO-003-12", "qty": 1.0, "unitCost": 17.1902471025585, "lineCost": 17.1902}, {"child": "MB-039", "qty": 0.00432, "unitCost": 96.0381355932203, "lineCost": 0.4149}, {"child": "MB-040", "qty": 0.00288, "unitCost": 114.00269541779, "lineCost": 0.3283}, {"child": "MB-101", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-102", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-103", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-104", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "MB-105", "qty": 0.00144, "unitCost": 265.0, "lineCost": 0.3816}, {"child": "PPH-001", "qty": 0.70992, "unitCost": 67.4507739938081, "lineCost": 47.8847}, {"child": "PPR-001", "qty": 1.92, "unitCost": 80.1845484377229, "lineCost": 153.9543}, {"child": "ST-EVO-003", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}]}, {"code": "EVO-101-12", "desc": "\u0637\u0642\u0645 \u062a\u062e\u0632\u064a\u0646 \u0627\u064a\u0641\u0648 \u0645\u0634\u0643\u0644", "conv": 12, "sp": 1688.0, "rawBoxCost": 537.4974, "boxCost": 725.6215, "profit": 962.3785, "profitPct": 57.01, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-EVO-104-12", "qty": 1.0, "unitCost": 14.9384138236597, "lineCost": 14.9384}, {"child": "MB-039", "qty": 0.01296, "unitCost": 96.0381355932203, "lineCost": 1.2447}, {"child": "MB-040", "qty": 0.00864, "unitCost": 114.00269541779, "lineCost": 0.985}, {"child": "MB-101", "qty": 0.00432, "unitCost": 265.0, "lineCost": 1.1448}, {"child": "MB-102", "qty": 0.00432, "unitCost": 265.0, "lineCost": 1.1448}, {"child": "MB-103", "qty": 0.00432, "unitCost": 265.0, "lineCost": 1.1448}, {"child": "MB-104", "qty": 0.00432, "unitCost": 265.0, "lineCost": 1.1448}, {"child": "MB-105", "qty": 0.00432, "unitCost": 265.0, "lineCost": 1.1448}, {"child": "PPH-001", "qty": 2.12976, "unitCost": 67.4507739938081, "lineCost": 143.654}, {"child": "PPR-001", "qty": 4.44, "unitCost": 80.1845484377229, "lineCost": 356.0194}, {"child": "ST-EVO-101", "qty": 12.0, "unitCost": 1.24432088959492, "lineCost": 14.9319}]}, {"code": "GLD-002-275", "desc": "\u0639\u0644\u0628\u0629  \u062c\u0648\u0644\u062f\u0646 \u0645\u0642\u0627\u0633 0.3 \u0644\u062a\u0631 \u062a\u0639\u0628\u0626\u0629 275 ", "conv": 550, "sp": 0, "rawBoxCost": 245.9417, "boxCost": 332.0213, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SGT-001", "qty": 1.0, "unitCost": 23.5291902071563, "lineCost": 23.5292}, {"child": "PPH-009", "qty": 6.05, "unitCost": 36.7623987644901, "lineCost": 222.4125}]}, {"code": "GLD-002-550", "desc": "\u0639\u0644\u0628\u0629  \u062c\u0648\u0644\u062f\u0646 \u0645\u0642\u0627\u0633 0.3 \u0644\u062a\u0631 ", "conv": 550, "sp": 0, "rawBoxCost": 473.4962, "boxCost": 639.2199, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GD-001", "qty": 1.0, "unitCost": 28.6711838306064, "lineCost": 28.6712}, {"child": "PPH-009", "qty": 12.1, "unitCost": 36.7623987644901, "lineCost": 444.825}]}, {"code": "GLD-003-350", "desc": "\u0639\u0644\u0628\u0629 \u062c\u0648\u0644\u062f\u0646 \u0645\u0642\u0627\u0633 0.5\u0644\u062a\u0631 ", "conv": 350, "sp": 0, "rawBoxCost": 569.0785, "boxCost": 768.256, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GD-001", "qty": 1.0, "unitCost": 28.6711838306064, "lineCost": 28.6712}, {"child": "PPH-009", "qty": 14.7, "unitCost": 36.7623987644901, "lineCost": 540.4073}]}, {"code": "GLD-004-220", "desc": "\u0639\u0644\u0628\u0629 \u062c\u0648\u0644\u062f\u0646 \u0645\u0642\u0627\u0633 1\u0644\u062a\u0631 ", "conv": 220, "sp": 0, "rawBoxCost": 522.0226, "boxCost": 704.7305, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GD-001", "qty": 1.0, "unitCost": 28.6711838306064, "lineCost": 28.6712}, {"child": "PPH-009", "qty": 13.42, "unitCost": 36.7623987644901, "lineCost": 493.3514}]}, {"code": "GLD-005-100", "desc": "\u0639\u0644\u0628\u0629 \u062c\u0648\u0644\u062f\u0646 \u0645\u0642\u0627\u0633 2\u0644\u062a\u0631 ", "conv": 100, "sp": 0, "rawBoxCost": 98.6712, "boxCost": 133.2061, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": ["SF-GT-05-1"], "bomLines": [{"child": "CE-GD-001", "qty": 1.0, "unitCost": 28.6711838306064, "lineCost": 28.6712}, {"child": "SF-GT-05-1", "qty": 100.0, "unitCost": null, "lineCost": 0.0}, {"child": "SF-GT-05-2", "qty": 100.0, "unitCost": 0.7, "lineCost": 70.0}]}, {"code": "GT-03-350", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.5\u0644\u062a\u0631 ", "conv": 350, "sp": 950.0, "rawBoxCost": 430.5321, "boxCost": 581.2183, "profit": 368.7817, "profitPct": 38.82, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GT-001", "qty": 1.0, "unitCost": 24.0432098765432, "lineCost": 24.0432}, {"child": "PPH-008", "qty": 12.6, "unitCost": 32.2610198789974, "lineCost": 406.4889}]}, {"code": "GT-04-220", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 1\u0644\u062a\u0631 ", "conv": 220, "sp": 950.0, "rawBoxCost": 421.499, "boxCost": 569.0237, "profit": 380.9763, "profitPct": 40.1, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GT-001", "qty": 1.0, "unitCost": 24.0432098765432, "lineCost": 24.0432}, {"child": "PPH-008", "qty": 12.32, "unitCost": 32.2610198789974, "lineCost": 397.4558}]}, {"code": "GT-11-325", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.20 \u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 \u0639\u062f\u062f 325 ", "conv": 650, "sp": 0, "rawBoxCost": 226.6415, "boxCost": 305.966, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SGT-001", "qty": 1.0, "unitCost": 23.5291902071563, "lineCost": 23.5292}, {"child": "PPH-009", "qty": 5.525, "unitCost": 36.7623987644901, "lineCost": 203.1123}]}, {"code": "GT-11-650", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.20 \u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 ", "conv": 650, "sp": 950.0, "rawBoxCost": 438.5687, "boxCost": 592.0677, "profit": 357.9323, "profitPct": 37.68, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PGT-001", "qty": 1.0, "unitCost": 32.3441550925926, "lineCost": 32.3442}, {"child": "PPH-009", "qty": 11.05, "unitCost": 36.7623987644901, "lineCost": 406.2245}]}, {"code": "GT-12-275", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.3 \u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 \u0639\u062f\u062f 275", "conv": 550, "sp": 0, "rawBoxCost": 254.7567, "boxCost": 343.9215, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PGT-001", "qty": 1.0, "unitCost": 32.3441550925926, "lineCost": 32.3442}, {"child": "PPH-009", "qty": 6.05, "unitCost": 36.7623987644901, "lineCost": 222.4125}]}, {"code": "GT-12-550", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.3 \u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636", "conv": 550, "sp": 950.0, "rawBoxCost": 477.1692, "boxCost": 644.1784, "profit": 305.8216, "profitPct": 32.19, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PGT-001", "qty": 1.0, "unitCost": 32.3441550925926, "lineCost": 32.3442}, {"child": "PPH-009", "qty": 12.1, "unitCost": 36.7623987644901, "lineCost": 444.825}]}, {"code": "GT-13-350", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 0.5\u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 ", "conv": 350, "sp": 950.0, "rawBoxCost": 572.7515, "boxCost": 773.2145, "profit": 176.7855, "profitPct": 18.61, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PGT-001", "qty": 1.0, "unitCost": 32.3441550925926, "lineCost": 32.3442}, {"child": "PPH-009", "qty": 14.7, "unitCost": 36.7623987644901, "lineCost": 540.4073}]}, {"code": "GT-14-220", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 1\u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 ", "conv": 220, "sp": 950.0, "rawBoxCost": 517.3946, "boxCost": 698.4827, "profit": 251.5173, "profitPct": 26.48, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-GT-001", "qty": 1.0, "unitCost": 24.0432098765432, "lineCost": 24.0432}, {"child": "PPH-009", "qty": 13.42, "unitCost": 36.7623987644901, "lineCost": 493.3514}]}, {"code": "GT-15-100", "desc": "\u0639\u0644\u0628\u0629 \u0645\u0642\u0627\u0633 2\u0644\u062a\u0631 \u0644\u0648\u0646 \u0627\u0628\u064a\u0636 ", "conv": 100, "sp": 850.0, "rawBoxCost": 403.6444, "boxCost": 544.9199, "profit": 305.0801, "profitPct": 35.89, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PGT-001", "qty": 1.0, "unitCost": 32.3441550925926, "lineCost": 32.3442}, {"child": "PPH-009", "qty": 10.1, "unitCost": 36.7623987644901, "lineCost": 371.3002}]}, {"code": "ICN-001-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 1400 \u0645\u0644", "conv": 24, "sp": 705.0, "rawBoxCost": 165.6985, "boxCost": 223.693, "profit": 481.307, "profitPct": 68.27, "hasBOM": true, "missing": ["SF-ICN-999-101", "SF-ICN-999-105"], "bomLines": [{"child": "CE-ICN-001-24", "qty": 1.0, "unitCost": 14.94, "lineCost": 14.94}, {"child": "MB-016", "qty": 0.0018, "unitCost": 118.069651741294, "lineCost": 0.2125}, {"child": "MB-040", "qty": 0.0018, "unitCost": 114.00269541779, "lineCost": 0.2052}, {"child": "MB-103", "qty": 0.0018, "unitCost": 265.0, "lineCost": 0.477}, {"child": "MB-104", "qty": 0.0018, "unitCost": 265.0, "lineCost": 0.477}, {"child": "PET-001", "qty": 1.8, "unitCost": 61.9588348579682, "lineCost": 111.5259}, {"child": "PPH-001", "qty": 0.24, "unitCost": 67.4507739938081, "lineCost": 16.1882}, {"child": "SF-ICN-999-101", "qty": 4.0, "unitCost": null, "lineCost": 0.0}, {"child": "SF-ICN-999-105", "qty": 4.0, "unitCost": null, "lineCost": 0.0}, {"child": "ST-ICN-001", "qty": 24.0, "unitCost": 0.903030303030303, "lineCost": 21.6727}]}, {"code": "ICN-002-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 800 \u0645\u0644", "conv": 24, "sp": 585.0, "rawBoxCost": 480.7037, "boxCost": 648.95, "profit": -63.95, "profitPct": -10.93, "hasBOM": true, "missing": ["SF-ICN-999-101", "SF-ICN-999-105"], "bomLines": [{"child": "CE-ICN-002-24", "qty": 3.0, "unitCost": 12.14, "lineCost": 36.42}, {"child": "MB-016", "qty": 0.0054, "unitCost": 118.069651741294, "lineCost": 0.6376}, {"child": "MB-040", "qty": 0.0054, "unitCost": 114.00269541779, "lineCost": 0.6156}, {"child": "MB-103", "qty": 0.0054, "unitCost": 265.0, "lineCost": 1.431}, {"child": "MB-104", "qty": 0.0054, "unitCost": 265.0, "lineCost": 1.431}, {"child": "PET-001", "qty": 1.32, "unitCost": 61.9588348579682, "lineCost": 81.7857}, {"child": "PPH-001", "qty": 0.72, "unitCost": 67.4507739938081, "lineCost": 48.5646}, {"child": "SF-ICN-002-00", "qty": 48.0, "unitCost": 5.25, "lineCost": 252.0}, {"child": "SF-ICN-999-101", "qty": 12.0, "unitCost": null, "lineCost": 0.0}, {"child": "SF-ICN-999-105", "qty": 12.0, "unitCost": null, "lineCost": 0.0}, {"child": "ST-ICN-002", "qty": 72.0, "unitCost": 0.803030303030303, "lineCost": 57.8182}]}, {"code": "ICN-005-48", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 1000\u0645\u0644 ", "conv": 48, "sp": 1680.0, "rawBoxCost": 345.3985, "boxCost": 466.288, "profit": 1213.712, "profitPct": 72.24, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-003", "qty": 1.0, "unitCost": 19.75, "lineCost": 19.75}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 4.08, "unitCost": 61.9588348579682, "lineCost": 252.792}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICN8-005", "qty": 48.0, "unitCost": 0.8, "lineCost": 38.4}]}, {"code": "ICN-008-48", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 1300\u0645\u0644 ", "conv": 48, "sp": 0, "rawBoxCost": 338.1985, "boxCost": 456.568, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-003", "qty": 1.0, "unitCost": 19.75, "lineCost": 19.75}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 4.08, "unitCost": 61.9588348579682, "lineCost": 252.792}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICNS-008", "qty": 48.0, "unitCost": 0.65, "lineCost": 31.2}]}, {"code": "ICN-103-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 400 \u0645\u0644 4*1", "conv": 24, "sp": 1440.0, "rawBoxCost": 272.0543, "boxCost": 367.2733, "profit": 1072.7267, "profitPct": 74.49, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-001", "qty": 1.0, "unitCost": 17.5, "lineCost": 17.5}, {"child": "MB-018", "qty": 0.0048, "unitCost": 100.427215189873, "lineCost": 0.4821}, {"child": "MB-024", "qty": 0.0048, "unitCost": 112.5, "lineCost": 0.54}, {"child": "MB-030", "qty": 0.0048, "unitCost": 223.736576152305, "lineCost": 1.0739}, {"child": "MB-039", "qty": 0.0048, "unitCost": 96.0381355932203, "lineCost": 0.461}, {"child": "MB-047", "qty": 0.0048, "unitCost": 114.0, "lineCost": 0.5472}, {"child": "MB-059", "qty": 0.0048, "unitCost": 220.0, "lineCost": 1.056}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.96, "unitCost": 67.4507739938081, "lineCost": 64.7527}, {"child": "ST-ICN8-001", "qty": 24.0, "unitCost": 0.3, "lineCost": 7.2}]}, {"code": "ICN-104-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 750 \u0645\u0644 2*1 ", "conv": 24, "sp": 1320.0, "rawBoxCost": 251.4979, "boxCost": 339.5222, "profit": 980.4778, "profitPct": 74.28, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-002", "qty": 1.0, "unitCost": 17.0, "lineCost": 17.0}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICN-004", "qty": 48.0, "unitCost": 0.45, "lineCost": 21.6}]}, {"code": "ICN-105-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 1000\u0645\u0644 2*1 ", "conv": 24, "sp": 0, "rawBoxCost": 345.3985, "boxCost": 466.288, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-003", "qty": 1.0, "unitCost": 19.75, "lineCost": 19.75}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 4.08, "unitCost": 61.9588348579682, "lineCost": 252.792}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICN8-005", "qty": 48.0, "unitCost": 0.8, "lineCost": 38.4}]}, {"code": "ICN-106-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 400 \u0645\u0644 4*1", "conv": 24, "sp": 0, "rawBoxCost": 272.0543, "boxCost": 367.2733, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-001", "qty": 1.0, "unitCost": 17.5, "lineCost": 17.5}, {"child": "MB-018", "qty": 0.0048, "unitCost": 100.427215189873, "lineCost": 0.4821}, {"child": "MB-024", "qty": 0.0048, "unitCost": 112.5, "lineCost": 0.54}, {"child": "MB-030", "qty": 0.0048, "unitCost": 223.736576152305, "lineCost": 1.0739}, {"child": "MB-039", "qty": 0.0048, "unitCost": 96.0381355932203, "lineCost": 0.461}, {"child": "MB-047", "qty": 0.0048, "unitCost": 114.0, "lineCost": 0.5472}, {"child": "MB-059", "qty": 0.0048, "unitCost": 220.0, "lineCost": 1.056}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.96, "unitCost": 67.4507739938081, "lineCost": 64.7527}, {"child": "ST-ICN8-001", "qty": 24.0, "unitCost": 0.3, "lineCost": 7.2}]}, {"code": "ICN-107-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 800\u0645\u0644 2*1 ", "conv": 24, "sp": 0, "rawBoxCost": 251.4979, "boxCost": 339.5222, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-002", "qty": 1.0, "unitCost": 17.0, "lineCost": 17.0}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICNS-007", "qty": 48.0, "unitCost": 0.45, "lineCost": 21.6}]}, {"code": "ICN-108-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 1300\u0645\u0644 2*1 ", "conv": 24, "sp": 0, "rawBoxCost": 338.1985, "boxCost": 456.568, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-003", "qty": 1.0, "unitCost": 19.75, "lineCost": 19.75}, {"child": "MB-018", "qty": 0.0024, "unitCost": 100.427215189873, "lineCost": 0.241}, {"child": "MB-024", "qty": 0.0024, "unitCost": 112.5, "lineCost": 0.27}, {"child": "MB-030", "qty": 0.0024, "unitCost": 223.736576152305, "lineCost": 0.537}, {"child": "MB-039", "qty": 0.0024, "unitCost": 96.0381355932203, "lineCost": 0.2305}, {"child": "MB-047", "qty": 0.0024, "unitCost": 114.0, "lineCost": 0.2736}, {"child": "MB-059", "qty": 0.0024, "unitCost": 220.0, "lineCost": 0.528}, {"child": "PET-001", "qty": 4.08, "unitCost": 61.9588348579682, "lineCost": 252.792}, {"child": "PPH-001", "qty": 0.48, "unitCost": 67.4507739938081, "lineCost": 32.3764}, {"child": "ST-ICNS-008", "qty": 48.0, "unitCost": 0.65, "lineCost": 31.2}]}, {"code": "ICN-203-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 400 \u0645\u0644 6*1", "conv": 24, "sp": 0, "rawBoxCost": 447.2316, "boxCost": 603.7627, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-101", "qty": 1.0, "unitCost": 21.0, "lineCost": 21.0}, {"child": "MB-018", "qty": 0.0072, "unitCost": 100.427215189873, "lineCost": 0.7231}, {"child": "MB-024", "qty": 0.0072, "unitCost": 112.5, "lineCost": 0.81}, {"child": "MB-030", "qty": 0.0072, "unitCost": 223.736576152305, "lineCost": 1.6109}, {"child": "MB-039", "qty": 0.0072, "unitCost": 96.0381355932203, "lineCost": 0.6915}, {"child": "MB-047", "qty": 0.0072, "unitCost": 114.0, "lineCost": 0.8208}, {"child": "MB-059", "qty": 0.0072, "unitCost": 220.0, "lineCost": 1.584}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 1.44, "unitCost": 67.4507739938081, "lineCost": 97.1291}, {"child": "ST-ICN8-102", "qty": 24.0, "unitCost": 2.3, "lineCost": 55.2}]}, {"code": "ICN-204-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 750 \u0645\u0644 3*1", "conv": 24, "sp": 0, "rawBoxCost": 371.7469, "boxCost": 501.8583, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-102", "qty": 1.0, "unitCost": 20.0, "lineCost": 20.0}, {"child": "MB-018", "qty": 0.0036, "unitCost": 100.427215189873, "lineCost": 0.3615}, {"child": "MB-024", "qty": 0.0036, "unitCost": 112.5, "lineCost": 0.405}, {"child": "MB-030", "qty": 0.0036, "unitCost": 223.736576152305, "lineCost": 0.8055}, {"child": "MB-039", "qty": 0.0036, "unitCost": 96.0381355932203, "lineCost": 0.3457}, {"child": "MB-047", "qty": 0.0036, "unitCost": 114.0, "lineCost": 0.4104}, {"child": "MB-059", "qty": 0.0036, "unitCost": 220.0, "lineCost": 0.792}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 0.72, "unitCost": 67.4507739938081, "lineCost": 48.5646}, {"child": "ST-ICN-004", "qty": 72.0, "unitCost": 0.45, "lineCost": 32.4}]}, {"code": "ICN-206-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 400 \u0645\u0644 6*1", "conv": 24, "sp": 0, "rawBoxCost": 447.2316, "boxCost": 603.7627, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-101", "qty": 1.0, "unitCost": 21.0, "lineCost": 21.0}, {"child": "MB-018", "qty": 0.0072, "unitCost": 100.427215189873, "lineCost": 0.7231}, {"child": "MB-024", "qty": 0.0072, "unitCost": 112.5, "lineCost": 0.81}, {"child": "MB-030", "qty": 0.0072, "unitCost": 223.736576152305, "lineCost": 1.6109}, {"child": "MB-039", "qty": 0.0072, "unitCost": 96.0381355932203, "lineCost": 0.6915}, {"child": "MB-047", "qty": 0.0072, "unitCost": 114.0, "lineCost": 0.8208}, {"child": "MB-059", "qty": 0.0072, "unitCost": 220.0, "lineCost": 1.584}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 1.44, "unitCost": 67.4507739938081, "lineCost": 97.1291}, {"child": "ST-ICN8-102", "qty": 24.0, "unitCost": 2.3, "lineCost": 55.2}]}, {"code": "ICN-207-24", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 800 \u0645\u0644 3*1", "conv": 24, "sp": 0, "rawBoxCost": 371.7469, "boxCost": 501.8583, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-102", "qty": 1.0, "unitCost": 20.0, "lineCost": 20.0}, {"child": "MB-018", "qty": 0.0036, "unitCost": 100.427215189873, "lineCost": 0.3615}, {"child": "MB-024", "qty": 0.0036, "unitCost": 112.5, "lineCost": 0.405}, {"child": "MB-030", "qty": 0.0036, "unitCost": 223.736576152305, "lineCost": 0.8055}, {"child": "MB-039", "qty": 0.0036, "unitCost": 96.0381355932203, "lineCost": 0.3457}, {"child": "MB-047", "qty": 0.0036, "unitCost": 114.0, "lineCost": 0.4104}, {"child": "MB-059", "qty": 0.0036, "unitCost": 220.0, "lineCost": 0.792}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 0.72, "unitCost": 67.4507739938081, "lineCost": 48.5646}, {"child": "ST-ICNS-007", "qty": 72.0, "unitCost": 0.45, "lineCost": 32.4}]}, {"code": "ICN-303-12", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 400 \u0645\u0644 8*1", "conv": 12, "sp": 0, "rawBoxCost": 300.8543, "boxCost": 406.1533, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-001", "qty": 1.0, "unitCost": 17.5, "lineCost": 17.5}, {"child": "MB-018", "qty": 0.0048, "unitCost": 100.427215189873, "lineCost": 0.4821}, {"child": "MB-024", "qty": 0.0048, "unitCost": 112.5, "lineCost": 0.54}, {"child": "MB-030", "qty": 0.0048, "unitCost": 223.736576152305, "lineCost": 1.0739}, {"child": "MB-039", "qty": 0.0048, "unitCost": 96.0381355932203, "lineCost": 0.461}, {"child": "MB-047", "qty": 0.0048, "unitCost": 114.0, "lineCost": 0.5472}, {"child": "MB-059", "qty": 0.0048, "unitCost": 220.0, "lineCost": 1.056}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.96, "unitCost": 67.4507739938081, "lineCost": 64.7527}, {"child": "ST-ICN8-101", "qty": 12.0, "unitCost": 3.0, "lineCost": 36.0}]}, {"code": "ICN-306-12", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 400 \u0645\u0644 8*1", "conv": 12, "sp": 0, "rawBoxCost": 300.8543, "boxCost": 406.1533, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-001", "qty": 1.0, "unitCost": 17.5, "lineCost": 17.5}, {"child": "MB-018", "qty": 0.0048, "unitCost": 100.427215189873, "lineCost": 0.4821}, {"child": "MB-024", "qty": 0.0048, "unitCost": 112.5, "lineCost": 0.54}, {"child": "MB-030", "qty": 0.0048, "unitCost": 223.736576152305, "lineCost": 1.0739}, {"child": "MB-039", "qty": 0.0048, "unitCost": 96.0381355932203, "lineCost": 0.461}, {"child": "MB-047", "qty": 0.0048, "unitCost": 114.0, "lineCost": 0.5472}, {"child": "MB-059", "qty": 0.0048, "unitCost": 220.0, "lineCost": 1.056}, {"child": "PET-001", "qty": 2.88, "unitCost": 61.9588348579682, "lineCost": 178.4414}, {"child": "PPH-001", "qty": 0.96, "unitCost": 67.4507739938081, "lineCost": 64.7527}, {"child": "ST-ICN8-101", "qty": 12.0, "unitCost": 3.0, "lineCost": 36.0}]}, {"code": "ICN-403-12", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 8 400 \u0645\u0644 12*1", "conv": 12, "sp": 2100.0, "rawBoxCost": 428.0316, "boxCost": 577.8427, "profit": 1522.1573, "profitPct": 72.48, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-101", "qty": 1.0, "unitCost": 21.0, "lineCost": 21.0}, {"child": "MB-018", "qty": 0.0072, "unitCost": 100.427215189873, "lineCost": 0.7231}, {"child": "MB-024", "qty": 0.0072, "unitCost": 112.5, "lineCost": 0.81}, {"child": "MB-030", "qty": 0.0072, "unitCost": 223.736576152305, "lineCost": 1.6109}, {"child": "MB-039", "qty": 0.0072, "unitCost": 96.0381355932203, "lineCost": 0.6915}, {"child": "MB-047", "qty": 0.0072, "unitCost": 114.0, "lineCost": 0.8208}, {"child": "MB-059", "qty": 0.0072, "unitCost": 220.0, "lineCost": 1.584}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 1.44, "unitCost": 67.4507739938081, "lineCost": 97.1291}, {"child": "ST-ICN8-101", "qty": 12.0, "unitCost": 3.0, "lineCost": 36.0}]}, {"code": "ICN-406-12", "desc": "\u0628\u0631\u0637\u0645\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0633\u0643\u0648\u064a\u0631 400 \u0645\u0644 12*1", "conv": 12, "sp": 0, "rawBoxCost": 428.0316, "boxCost": 577.8427, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN8-101", "qty": 1.0, "unitCost": 21.0, "lineCost": 21.0}, {"child": "MB-018", "qty": 0.0072, "unitCost": 100.427215189873, "lineCost": 0.7231}, {"child": "MB-024", "qty": 0.0072, "unitCost": 112.5, "lineCost": 0.81}, {"child": "MB-030", "qty": 0.0072, "unitCost": 223.736576152305, "lineCost": 1.6109}, {"child": "MB-039", "qty": 0.0072, "unitCost": 96.0381355932203, "lineCost": 0.6915}, {"child": "MB-047", "qty": 0.0072, "unitCost": 114.0, "lineCost": 0.8208}, {"child": "MB-059", "qty": 0.0072, "unitCost": 220.0, "lineCost": 1.584}, {"child": "PET-001", "qty": 4.32, "unitCost": 61.9588348579682, "lineCost": 267.6622}, {"child": "PPH-001", "qty": 1.44, "unitCost": 67.4507739938081, "lineCost": 97.1291}, {"child": "ST-ICN8-101", "qty": 12.0, "unitCost": 3.0, "lineCost": 36.0}]}, {"code": "ICN-501-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0645\u0631\u0637\u0628\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0645\u0631\u0628\u06397*1", "conv": 12, "sp": 2580.0, "rawBoxCost": 650.9946, "boxCost": 878.8427, "profit": 1701.1573, "profitPct": 65.94, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN-501-12", "qty": 1.0, "unitCost": 22.7, "lineCost": 22.7}, {"child": "MB-018", "qty": 0.01998, "unitCost": 100.427215189873, "lineCost": 2.0065}, {"child": "MB-024", "qty": 0.01998, "unitCost": 112.5, "lineCost": 2.2478}, {"child": "MB-030", "qty": 0.01998, "unitCost": 223.736576152305, "lineCost": 4.4703}, {"child": "MB-039", "qty": 0.01998, "unitCost": 96.0381355932203, "lineCost": 1.9188}, {"child": "MB-047", "qty": 0.01998, "unitCost": 114.0, "lineCost": 2.2777}, {"child": "MB-059", "qty": 0.01998, "unitCost": 220.0, "lineCost": 4.3956}, {"child": "PET-001", "qty": 4.26, "unitCost": 61.9588348579682, "lineCost": 263.9446}, {"child": "PPH-001", "qty": 3.996, "unitCost": 67.4507739938081, "lineCost": 269.5333}, {"child": "PSH-ICN-501", "qty": 0.5, "unitCost": 95.0, "lineCost": 47.5}, {"child": "ST-ICN-501-01", "qty": 12.0, "unitCost": 2.5, "lineCost": 30.0}]}, {"code": "ICN-601-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u0645\u0631\u0637\u0628\u0627\u0646 \u0627\u064a\u0643\u0648\u0646 \u0645\u0636\u0644\u0639 7*1", "conv": 12, "sp": 0, "rawBoxCost": 650.9946, "boxCost": 878.8427, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ICN-501-12", "qty": 1.0, "unitCost": 22.7, "lineCost": 22.7}, {"child": "MB-018", "qty": 0.01998, "unitCost": 100.427215189873, "lineCost": 2.0065}, {"child": "MB-024", "qty": 0.01998, "unitCost": 112.5, "lineCost": 2.2478}, {"child": "MB-030", "qty": 0.01998, "unitCost": 223.736576152305, "lineCost": 4.4703}, {"child": "MB-039", "qty": 0.01998, "unitCost": 96.0381355932203, "lineCost": 1.9188}, {"child": "MB-047", "qty": 0.01998, "unitCost": 114.0, "lineCost": 2.2777}, {"child": "MB-059", "qty": 0.01998, "unitCost": 220.0, "lineCost": 4.3956}, {"child": "PET-001", "qty": 4.26, "unitCost": 61.9588348579682, "lineCost": 263.9446}, {"child": "PPH-001", "qty": 3.996, "unitCost": 67.4507739938081, "lineCost": 269.5333}, {"child": "PSH-ICN-501", "qty": 0.5, "unitCost": 95.0, "lineCost": 47.5}, {"child": "ST-ICN-501-01", "qty": 12.0, "unitCost": 2.5, "lineCost": 30.0}]}, {"code": "KEP-001-12", "desc": "\u062d\u0627\u0641\u0638\u0629 \u0645\u062e\u0628\u0648\u0632\u0627\u062a \u0643\u064a\u0628\u064a\u0631", "conv": 12, "sp": 1375.0, "rawBoxCost": 443.1128, "boxCost": 598.2023, "profit": 776.7977, "profitPct": 56.49, "hasBOM": true, "missing": [], "bomLines": [{"child": "MB-046", "qty": 0.01422, "unitCost": 112.599364069952, "lineCost": 1.6012}, {"child": "PPH-001", "qty": 0.084, "unitCost": 67.4507739938081, "lineCost": 5.6659}, {"child": "PPK-001", "qty": 1.54, "unitCost": 75.3691770186335, "lineCost": 116.0685}, {"child": "PPK-101", "qty": 1.1, "unitCost": 51.4602533485619, "lineCost": 56.6063}, {"child": "PPR-001", "qty": 2.7, "unitCost": 80.1845484377229, "lineCost": 216.4983}, {"child": "ST-KEP-001", "qty": 12.0, "unitCost": 0.84925735797399, "lineCost": 10.1911}, {"child": "CE-KEP-001-12", "qty": 1.0, "unitCost": 26.6338831045906, "lineCost": 26.6339}, {"child": "MB-013", "qty": 0.0104, "unitCost": 190.306387225549, "lineCost": 1.9792}, {"child": "MB-016", "qty": 0.01422, "unitCost": 118.069651741294, "lineCost": 1.679}, {"child": "MB-019", "qty": 0.01422, "unitCost": 125.283018867925, "lineCost": 1.7815}, {"child": "MB-038", "qty": 0.00042, "unitCost": 185.610749185668, "lineCost": 0.078}, {"child": "MB-040", "qty": 0.01422, "unitCost": 114.00269541779, "lineCost": 1.6211}, {"child": "MB-045", "qty": 0.01422, "unitCost": 190.492957746479, "lineCost": 2.7088}]}, {"code": "KEP-001-36", "desc": "\u062d\u0627\u0641\u0638\u0629 \u0645\u062e\u0628\u0648\u0632\u0627\u062a \u0643\u064a\u0628\u064a\u0631 \u0639\u062f\u062f 36", "conv": 36, "sp": 0, "rawBoxCost": 1295.9964, "boxCost": 1749.5951, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-KEP-001-36", "qty": 1.0, "unitCost": 46.56, "lineCost": 46.56}, {"child": "MB-013", "qty": 0.0312, "unitCost": 190.306387225549, "lineCost": 5.9376}, {"child": "MB-016", "qty": 0.04266, "unitCost": 118.069651741294, "lineCost": 5.0369}, {"child": "MB-019", "qty": 0.04266, "unitCost": 125.283018867925, "lineCost": 5.3446}, {"child": "MB-038", "qty": 0.00126, "unitCost": 185.610749185668, "lineCost": 0.2339}, {"child": "MB-040", "qty": 0.04266, "unitCost": 114.00269541779, "lineCost": 4.8634}, {"child": "MB-045", "qty": 0.04266, "unitCost": 190.492957746479, "lineCost": 8.1264}, {"child": "MB-046", "qty": 0.04266, "unitCost": 112.599364069952, "lineCost": 4.8035}, {"child": "PPH-001", "qty": 0.252, "unitCost": 67.4507739938081, "lineCost": 16.9976}, {"child": "PPK-001", "qty": 4.62, "unitCost": 75.3691770186335, "lineCost": 348.2056}, {"child": "PPK-101", "qty": 3.3, "unitCost": 51.4602533485619, "lineCost": 169.8188}, {"child": "PPR-001", "qty": 8.1, "unitCost": 80.1845484377229, "lineCost": 649.4948}, {"child": "ST-KEP-001", "qty": 36.0, "unitCost": 0.84925735797399, "lineCost": 30.5733}]}, {"code": "LST-001-12", "desc": "\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0633\u0645\u0627\u0631\u062a", "conv": 12, "sp": 906.0, "rawBoxCost": 312.0189, "boxCost": 421.2255, "profit": 484.7745, "profitPct": 53.51, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SMT-001-12", "qty": 1.0, "unitCost": 11.8783429672447, "lineCost": 11.8783}, {"child": "MB-032", "qty": 0.00678, "unitCost": 190.363636363636, "lineCost": 1.2907}, {"child": "MB-035", "qty": 0.00678, "unitCost": 186.397467572576, "lineCost": 1.2638}, {"child": "MB-036", "qty": 0.00678, "unitCost": 188.125, "lineCost": 1.2755}, {"child": "MB-043", "qty": 0.00678, "unitCost": 170.799180327869, "lineCost": 1.158}, {"child": "MB-044", "qty": 0.00678, "unitCost": 188.547619047619, "lineCost": 1.2784}, {"child": "MB-055", "qty": 0.00678, "unitCost": 214.316753926702, "lineCost": 1.4531}, {"child": "PPK-001", "qty": 1.2, "unitCost": 75.3691770186335, "lineCost": 90.443}, {"child": "PPR-001", "qty": 1.74, "unitCost": 80.1845484377229, "lineCost": 139.5211}, {"child": "RUB-001", "qty": 0.312, "unitCost": 182.875, "lineCost": 57.057}, {"child": "ST-LST-001-01", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}]}, {"code": "LST-101-12", "desc": "\u0645\u062c\u0645\u0648\u0639\u0629 \u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0633\u0645\u0627\u0631\u062a", "conv": 12, "sp": 1125.0, "rawBoxCost": 378.5873, "boxCost": 511.0929, "profit": 613.9071, "profitPct": 54.57, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SMT-001-12", "qty": 1.0, "unitCost": 11.8783429672447, "lineCost": 11.8783}, {"child": "MB-032", "qty": 0.00738, "unitCost": 190.363636363636, "lineCost": 1.4049}, {"child": "MB-035", "qty": 0.00882, "unitCost": 186.397467572576, "lineCost": 1.644}, {"child": "MB-036", "qty": 0.0081, "unitCost": 188.125, "lineCost": 1.5238}, {"child": "MB-043", "qty": 0.0081, "unitCost": 170.799180327869, "lineCost": 1.3835}, {"child": "MB-044", "qty": 0.0081, "unitCost": 188.547619047619, "lineCost": 1.5272}, {"child": "MB-050", "qty": 0.00072, "unitCost": 114.0, "lineCost": 0.0821}, {"child": "MB-055", "qty": 0.0081, "unitCost": 214.316753926702, "lineCost": 1.736}, {"child": "PET-001", "qty": 0.72, "unitCost": 61.9588348579682, "lineCost": 44.6104}, {"child": "PPH-001", "qty": 0.168, "unitCost": 67.4507739938081, "lineCost": 11.3317}, {"child": "PPK-001", "qty": 1.32, "unitCost": 75.3691770186335, "lineCost": 99.4873}, {"child": "PPR-001", "qty": 1.74, "unitCost": 80.1845484377229, "lineCost": 139.5211}, {"child": "RUB-001", "qty": 0.312, "unitCost": 182.875, "lineCost": 57.057}, {"child": "ST-LST-001-01", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}]}, {"code": "LUX-001-12", "desc": "\u0631\u0641 \u0644\u0648\u0643\u0633", "conv": 12, "sp": 1725.0, "rawBoxCost": 559.7637, "boxCost": 755.681, "profit": 969.319, "profitPct": 56.19, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-LUX-001-12", "qty": 1.0, "unitCost": 28.5153029356652, "lineCost": 28.5153}, {"child": "CI-LUX-001-01", "qty": 12.0, "unitCost": 8.45486974734913, "lineCost": 101.4584}, {"child": "MB-014", "qty": 0.01068, "unitCost": 105.0, "lineCost": 1.1214}, {"child": "MB-015", "qty": 0.0252, "unitCost": 119.240506329114, "lineCost": 3.0049}, {"child": "MB-020", "qty": 0.01068, "unitCost": 121.964285714286, "lineCost": 1.3026}, {"child": "MB-026", "qty": 0.0252, "unitCost": 316.0, "lineCost": 7.9632}, {"child": "MB-030", "qty": 0.01068, "unitCost": 223.736576152305, "lineCost": 2.3895}, {"child": "MB-037", "qty": 0.01068, "unitCost": 109.956331877729, "lineCost": 1.1743}, {"child": "MB-038", "qty": 0.0504, "unitCost": 185.610749185668, "lineCost": 9.3548}, {"child": "MB-040", "qty": 0.03588, "unitCost": 114.00269541779, "lineCost": 4.0904}, {"child": "MB-046", "qty": 0.03588, "unitCost": 112.599364069952, "lineCost": 4.0401}, {"child": "PPH-005", "qty": 7.11192, "unitCost": 55.5896038716616, "lineCost": 395.3488}]}, {"code": "MIX-001-01", "desc": "\u062f\u0648\u0631\u0627\u0642 \u0645\u064a\u0643\u0633\u0631 \u0628\u0627\u0644\u0643\u0631\u062a\u0648\u0646\u0629", "conv": 12, "sp": 840.0, "rawBoxCost": 379.4423, "boxCost": 512.2471, "profit": 327.7529, "profitPct": 39.02, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-MIX-001-01", "qty": 12.0, "unitCost": 5.5, "lineCost": 66.0}, {"child": "MB-016", "qty": 0.0044, "unitCost": 118.069651741294, "lineCost": 0.5195}, {"child": "MB-039", "qty": 0.03252, "unitCost": 96.0381355932203, "lineCost": 3.1232}, {"child": "MB-040", "qty": 0.0044, "unitCost": 114.00269541779, "lineCost": 0.5016}, {"child": "MB-046", "qty": 0.0044, "unitCost": 112.599364069952, "lineCost": 0.4954}, {"child": "MB-047", "qty": 0.0044, "unitCost": 114.0, "lineCost": 0.5016}, {"child": "MB-049", "qty": 0.0044, "unitCost": 110.0, "lineCost": 0.484}, {"child": "MB-050", "qty": 0.0044, "unitCost": 114.0, "lineCost": 0.5016}, {"child": "PPH-001", "qty": 0.5706, "unitCost": 67.4507739938081, "lineCost": 38.4874}, {"child": "PPK-001", "qty": 0.84972, "unitCost": 75.3691770186335, "lineCost": 64.0427}, {"child": "PPR-001", "qty": 2.4, "unitCost": 80.1845484377229, "lineCost": 192.4429}, {"child": "SF-MIX-001-08-99", "qty": 12.0, "unitCost": 0.247, "lineCost": 2.964}, {"child": "SF-MIX-001-09-99", "qty": 12.0, "unitCost": 0.115, "lineCost": 1.38}, {"child": "ST-MIX-001", "qty": 12.0, "unitCost": 0.666536960234204, "lineCost": 7.9984}]}, {"code": "MIX-001-12", "desc": "\u062f\u0648\u0631\u0627\u0642 \u0645\u064a\u0643\u0633\u0631", "conv": 12, "sp": 1050.0, "rawBoxCost": 331.2591, "boxCost": 447.1998, "profit": 602.8002, "profitPct": 57.41, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-MIX-001-12", "qty": 1.0, "unitCost": 17.8168245765145, "lineCost": 17.8168}, {"child": "MB-016", "qty": 0.0044, "unitCost": 118.069651741294, "lineCost": 0.5195}, {"child": "MB-039", "qty": 0.03252, "unitCost": 96.0381355932203, "lineCost": 3.1232}, {"child": "MB-040", "qty": 0.0044, "unitCost": 114.00269541779, "lineCost": 0.5016}, {"child": "MB-046", "qty": 0.0044, "unitCost": 112.599364069952, "lineCost": 0.4954}, {"child": "MB-047", "qty": 0.0044, "unitCost": 114.0, "lineCost": 0.5016}, {"child": "MB-049", "qty": 0.0044, "unitCost": 110.0, "lineCost": 0.484}, {"child": "MB-050", "qty": 0.0044, "unitCost": 114.0, "lineCost": 0.5016}, {"child": "PPH-001", "qty": 0.5706, "unitCost": 67.4507739938081, "lineCost": 38.4874}, {"child": "PPK-001", "qty": 0.84972, "unitCost": 75.3691770186335, "lineCost": 64.0427}, {"child": "PPR-001", "qty": 2.4, "unitCost": 80.1845484377229, "lineCost": 192.4429}, {"child": "SF-MIX-001-08-99", "qty": 12.0, "unitCost": 0.247, "lineCost": 2.964}, {"child": "SF-MIX-001-09-99", "qty": 12.0, "unitCost": 0.115, "lineCost": 1.38}, {"child": "ST-MIX-001", "qty": 12.0, "unitCost": 0.666536960234204, "lineCost": 7.9984}]}, {"code": "PAP-001-12", "desc": "\u0637\u0628\u0642 \u0645\u0642\u0628\u0644\u0627\u062a \u0648 \u062a\u0633\u0627\u0644\u064a/5 \u0639\u064a\u0646", "conv": 60, "sp": 1594.0, "rawBoxCost": 487.0311, "boxCost": 657.492, "profit": 936.508, "profitPct": 58.75, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PAP-001-60", "qty": 1.0, "unitCost": 9.4695514092894, "lineCost": 9.4696}, {"child": "MB-016", "qty": 0.048, "unitCost": 118.069651741294, "lineCost": 5.6673}, {"child": "MB-017", "qty": 0.048, "unitCost": 95.0, "lineCost": 4.56}, {"child": "MB-021", "qty": 0.048, "unitCost": 126.344086021505, "lineCost": 6.0645}, {"child": "MB-025", "qty": 0.048, "unitCost": 97.5, "lineCost": 4.68}, {"child": "MB-040", "qty": 0.048, "unitCost": 114.00269541779, "lineCost": 5.4721}, {"child": "MB-041", "qty": 0.048, "unitCost": 170.486111111111, "lineCost": 8.1833}, {"child": "PPH-005", "qty": 7.86, "unitCost": 55.5896038716616, "lineCost": 436.9343}, {"child": "ST-PAP-001", "qty": 60.0, "unitCost": 0.1, "lineCost": 6.0}]}, {"code": "PAP-101- 12", "desc": "\u0637\u0642\u0645 \u0637\u0628\u0642 \u062a\u0633\u0627\u0644\u064a \u0639\u062f\u062f 12 \u0637\u0642\u0645 ", "conv": 12, "sp": 0, "rawBoxCost": 293.6066, "boxCost": 396.3689, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PAP-001-60", "qty": 1.0, "unitCost": 9.4695514092894, "lineCost": 9.4696}, {"child": "MB-016", "qty": 0.0288, "unitCost": 118.069651741294, "lineCost": 3.4004}, {"child": "MB-017", "qty": 0.0288, "unitCost": 95.0, "lineCost": 2.736}, {"child": "MB-021", "qty": 0.0288, "unitCost": 126.344086021505, "lineCost": 3.6387}, {"child": "MB-025", "qty": 0.0288, "unitCost": 97.5, "lineCost": 2.808}, {"child": "MB-040", "qty": 0.0288, "unitCost": 114.00269541779, "lineCost": 3.2833}, {"child": "MB-041", "qty": 0.0288, "unitCost": 170.486111111111, "lineCost": 4.91}, {"child": "PPH-005", "qty": 4.716, "unitCost": 55.5896038716616, "lineCost": 262.1606}, {"child": "ST-PAP-101", "qty": 12.0, "unitCost": 0.1, "lineCost": 1.2}]}, {"code": "PCR-001-06", "desc": "\u0631\u0643\u0646\u0629 4 \u062f\u0648\u0631", "conv": 6, "sp": 1281.0, "rawBoxCost": 378.6049, "boxCost": 511.1166, "profit": 769.8834, "profitPct": 60.1, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PCR-001-06", "qty": 1.0, "unitCost": 23.4489858793325, "lineCost": 23.449}, {"child": "MB-013", "qty": 0.00984, "unitCost": 190.306387225549, "lineCost": 1.8726}, {"child": "MB-015", "qty": 0.024, "unitCost": 119.240506329114, "lineCost": 2.8618}, {"child": "MB-018", "qty": 0.0051, "unitCost": 100.427215189873, "lineCost": 0.5122}, {"child": "MB-020", "qty": 0.01002, "unitCost": 121.964285714286, "lineCost": 1.2221}, {"child": "MB-030", "qty": 0.02436, "unitCost": 223.736576152305, "lineCost": 5.4502}, {"child": "MB-037", "qty": 0.0348, "unitCost": 109.956331877729, "lineCost": 3.8265}, {"child": "MB-038", "qty": 0.02418, "unitCost": 185.610749185668, "lineCost": 4.4881}, {"child": "MB-040", "qty": 0.024, "unitCost": 114.00269541779, "lineCost": 2.7361}, {"child": "MB-045", "qty": 0.024, "unitCost": 190.492957746479, "lineCost": 4.5718}, {"child": "PPH-003", "qty": 0.40572, "unitCost": 32.0, "lineCost": 12.983}, {"child": "PPH-004", "qty": 0.492, "unitCost": 43.7207771717081, "lineCost": 21.5106}, {"child": "PPH-005", "qty": 0.62298, "unitCost": 55.5896038716616, "lineCost": 34.6312}, {"child": "PPH-007", "qty": 0.24108, "unitCost": 35.1771189071717, "lineCost": 8.4805}, {"child": "PPK-101", "qty": 4.8, "unitCost": 51.4602533485619, "lineCost": 247.0092}, {"child": "ST-PCR-001", "qty": 6.0, "unitCost": 0.5, "lineCost": 3.0}]}, {"code": "PCR-002-12", "desc": "\u0631\u0643\u0646\u0629 2 \u062f\u0648\u0631", "conv": 12, "sp": 1400.0, "rawBoxCost": 366.8428, "boxCost": 495.2378, "profit": 904.7622, "profitPct": 64.63, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PCR-002-12", "qty": 1.0, "unitCost": 25.2789808917197, "lineCost": 25.279}, {"child": "MB-013", "qty": 0.00864, "unitCost": 190.306387225549, "lineCost": 1.6442}, {"child": "MB-015", "qty": 0.024, "unitCost": 119.240506329114, "lineCost": 2.8618}, {"child": "MB-018", "qty": 0.00468, "unitCost": 100.427215189873, "lineCost": 0.47}, {"child": "MB-020", "qty": 0.009, "unitCost": 121.964285714286, "lineCost": 1.0977}, {"child": "MB-030", "qty": 0.02472, "unitCost": 223.736576152305, "lineCost": 5.5308}, {"child": "MB-037", "qty": 0.02868, "unitCost": 109.956331877729, "lineCost": 3.1535}, {"child": "MB-038", "qty": 0.02436, "unitCost": 185.610749185668, "lineCost": 4.5215}, {"child": "MB-040", "qty": 0.024, "unitCost": 114.00269541779, "lineCost": 2.7361}, {"child": "MB-045", "qty": 0.024, "unitCost": 190.492957746479, "lineCost": 4.5718}, {"child": "PPH-003", "qty": 0.13524, "unitCost": 32.0, "lineCost": 4.3277}, {"child": "PPH-004", "qty": 0.432, "unitCost": 43.7207771717081, "lineCost": 18.8874}, {"child": "PPH-005", "qty": 0.56316, "unitCost": 55.5896038716616, "lineCost": 31.3058}, {"child": "PPH-007", "qty": 0.21168, "unitCost": 35.1771189071717, "lineCost": 7.4463}, {"child": "PPK-101", "qty": 4.8, "unitCost": 51.4602533485619, "lineCost": 247.0092}, {"child": "ST-PCR-002", "qty": 12.0, "unitCost": 0.5, "lineCost": 6.0}]}, {"code": "PCU-001-12", "desc": "\u0642\u0637\u0627\u0639\u0629 \u062e\u0636\u0627\u0631", "conv": 12, "sp": 950.0, "rawBoxCost": 311.0105, "boxCost": 419.8642, "profit": 530.1358, "profitPct": 55.8, "hasBOM": true, "missing": ["SF-PST-001-01-36"], "bomLines": [{"child": "CE-PCU-001-12", "qty": 1.0, "unitCost": 18.1941690053009, "lineCost": 18.1942}, {"child": "MB-021", "qty": 0.012, "unitCost": 126.344086021505, "lineCost": 1.5161}, {"child": "MB-039", "qty": 0.063, "unitCost": 96.0381355932203, "lineCost": 6.0504}, {"child": "MB-040", "qty": 0.012, "unitCost": 114.00269541779, "lineCost": 1.368}, {"child": "MB-046", "qty": 0.012, "unitCost": 112.599364069952, "lineCost": 1.3512}, {"child": "MB-048", "qty": 0.012, "unitCost": 110.0, "lineCost": 1.32}, {"child": "MB-050", "qty": 0.012, "unitCost": 114.0, "lineCost": 1.368}, {"child": "PPH-005", "qty": 2.0, "unitCost": 55.5896038716616, "lineCost": 111.1792}, {"child": "PPK-001", "qty": 2.1, "unitCost": 75.3691770186335, "lineCost": 158.2753}, {"child": "SF-PST-001-01-36", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "ST-PCU-001", "qty": 12.0, "unitCost": 0.865671641791045, "lineCost": 10.3881}]}, {"code": "PDO-001-12", "desc": "\u0645\u0646\u0638\u0645 \u0627\u0637\u0628\u0627\u0642 \u0628\u064a\u0631\u0641\u0643\u062a", "conv": 12, "sp": 1650.0, "rawBoxCost": 475.3741, "boxCost": 641.755, "profit": 1008.245, "profitPct": 61.11, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PDO-001-12", "qty": 1.0, "unitCost": 34.9094231484927, "lineCost": 34.9094}, {"child": "MB-026", "qty": 0.042, "unitCost": 316.0, "lineCost": 13.272}, {"child": "MB-030", "qty": 0.036, "unitCost": 223.736576152305, "lineCost": 8.0545}, {"child": "MB-037", "qty": 0.057, "unitCost": 109.956331877729, "lineCost": 6.2675}, {"child": "MB-038", "qty": 0.036, "unitCost": 185.610749185668, "lineCost": 6.682}, {"child": "MB-040", "qty": 0.036, "unitCost": 114.00269541779, "lineCost": 4.1041}, {"child": "MB-046", "qty": 0.036, "unitCost": 112.599364069952, "lineCost": 4.0536}, {"child": "PPK-101", "qty": 7.625, "unitCost": 51.4602533485619, "lineCost": 392.3844}, {"child": "ST-PDO-001", "qty": 12.0, "unitCost": 0.470548317046688, "lineCost": 5.6466}]}, {"code": "PDP-001-24", "desc": "\u0637\u0628\u0642 \u0645\u0636\u0644\u0639", "conv": 24, "sp": 725.0, "rawBoxCost": 223.3636, "boxCost": 301.5409, "profit": 423.4591, "profitPct": 58.41, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PDP-001-24", "qty": 1.0, "unitCost": 10.75, "lineCost": 10.75}, {"child": "MB-016", "qty": 0.018, "unitCost": 118.069651741294, "lineCost": 2.1253}, {"child": "MB-021", "qty": 0.018, "unitCost": 126.344086021505, "lineCost": 2.2742}, {"child": "MB-040", "qty": 0.018, "unitCost": 114.00269541779, "lineCost": 2.052}, {"child": "MB-048", "qty": 0.018, "unitCost": 110.0, "lineCost": 1.98}, {"child": "MB-050", "qty": 0.018, "unitCost": 114.0, "lineCost": 2.052}, {"child": "MB-051", "qty": 0.018, "unitCost": 111.527777777778, "lineCost": 2.0075}, {"child": "PPH-005", "qty": 3.6, "unitCost": 55.5896038716616, "lineCost": 200.1226}]}, {"code": "PJS-001-12", "desc": "\u0639\u0635\u0627\u0631\u0629 \u0645\u0648\u0627\u0644\u062d", "conv": 12, "sp": 294.0, "rawBoxCost": 62.6446, "boxCost": 84.5702, "profit": 209.4298, "profitPct": 71.23, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PJS-001-12", "qty": 1.0, "unitCost": 9.02827045274027, "lineCost": 9.0283}, {"child": "DPR-001", "qty": 0.648, "unitCost": 38.4615, "lineCost": 24.9231}, {"child": "MB-016", "qty": 0.0021, "unitCost": 118.069651741294, "lineCost": 0.2479}, {"child": "MB-017", "qty": 0.0021, "unitCost": 95.0, "lineCost": 0.1995}, {"child": "MB-027", "qty": 0.0021, "unitCost": 130.0, "lineCost": 0.273}, {"child": "MB-028", "qty": 0.0021, "unitCost": 130.0, "lineCost": 0.273}, {"child": "MB-040", "qty": 0.0021, "unitCost": 114.00269541779, "lineCost": 0.2394}, {"child": "MB-050", "qty": 0.0021, "unitCost": 114.0, "lineCost": 0.2394}, {"child": "PPH-005", "qty": 0.4074, "unitCost": 55.5896038716616, "lineCost": 22.6472}, {"child": "ST-PJS-001", "qty": 12.0, "unitCost": 0.381150159744409, "lineCost": 4.5738}]}, {"code": "PLN-001-12", "desc": "\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0643\u064a\u062f\u0632", "conv": 12, "sp": 263.0, "rawBoxCost": 139.4533, "boxCost": 188.262, "profit": 74.738, "profitPct": 28.42, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PLN-001-12", "qty": 1.0, "unitCost": 6.0, "lineCost": 6.0}, {"child": "MB-032", "qty": 0.045, "unitCost": 190.363636363636, "lineCost": 8.5664}, {"child": "MB-035", "qty": 0.045, "unitCost": 186.397467572576, "lineCost": 8.3879}, {"child": "MB-036", "qty": 0.045, "unitCost": 188.125, "lineCost": 8.4656}, {"child": "MB-037", "qty": 0.00096, "unitCost": 109.956331877729, "lineCost": 0.1056}, {"child": "MB-043", "qty": 0.045, "unitCost": 170.799180327869, "lineCost": 7.686}, {"child": "MB-044", "qty": 0.045, "unitCost": 188.547619047619, "lineCost": 8.4846}, {"child": "MB-055", "qty": 0.045, "unitCost": 214.316753926702, "lineCost": 9.6443}, {"child": "PPR-001", "qty": 0.9, "unitCost": 80.1845484377229, "lineCost": 72.1661}, {"child": "PPS-101", "qty": 0.03, "unitCost": 71.5607979734009, "lineCost": 2.1468}, {"child": "ST-PLN-002", "qty": 12.0, "unitCost": 0.65, "lineCost": 7.8}]}, {"code": "PLN-002-12", "desc": "\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0647\u064a\u0631\u0648", "conv": 12, "sp": 563.0, "rawBoxCost": 3036.3444, "boxCost": 4099.0649, "profit": -3536.0649, "profitPct": -628.08, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PLN-002-12", "qty": 1.0, "unitCost": 10.4484439402534, "lineCost": 10.4484}, {"child": "MB-032", "qty": 0.0459, "unitCost": 190.363636363636, "lineCost": 8.7377}, {"child": "MB-035", "qty": 0.0459, "unitCost": 186.397467572576, "lineCost": 8.5556}, {"child": "MB-036", "qty": 0.0459, "unitCost": 188.125, "lineCost": 8.6349}, {"child": "MB-039", "qty": 0.0216, "unitCost": 96.0381355932203, "lineCost": 2.0744}, {"child": "MB-043", "qty": 0.0459, "unitCost": 170.799180327869, "lineCost": 7.8397}, {"child": "MB-044", "qty": 0.0459, "unitCost": 188.547619047619, "lineCost": 8.6543}, {"child": "MB-055", "qty": 0.0459, "unitCost": 214.316753926702, "lineCost": 9.8371}, {"child": "PPH-001", "qty": 0.72, "unitCost": 67.4507739938081, "lineCost": 48.5646}, {"child": "PPH-005", "qty": 0.33, "unitCost": 55.5896038716616, "lineCost": 18.3446}, {"child": "PPR-001", "qty": 1.38, "unitCost": 80.1845484377229, "lineCost": 110.6547}, {"child": "ST-PLN-001", "qty": 12.0, "unitCost": 0.890222222222222, "lineCost": 10.6827}]}, {"code": "PLT-001-12", "desc": "\u0637\u0628\u0642 \u0628\u0644\u0627\u064a\u062a \u0628\u0644\u0633 ", "conv": 12, "sp": 1350.0, "rawBoxCost": 397.5897, "boxCost": 536.7461, "profit": 813.2539, "profitPct": 60.24, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PLT-001-12", "qty": 1.0, "unitCost": 21.67, "lineCost": 21.67}, {"child": "CI-PLT-001-01", "qty": 12.0, "unitCost": 7.0, "lineCost": 84.0}, {"child": "MB-039", "qty": 0.003, "unitCost": 96.0381355932203, "lineCost": 0.2881}, {"child": "MB-052", "qty": 0.006, "unitCost": 100.0, "lineCost": 0.6}, {"child": "PPH-001", "qty": 0.3, "unitCost": 67.4507739938081, "lineCost": 20.2352}, {"child": "PPS-001", "qty": 3.12756, "unitCost": 86.5839025839026, "lineCost": 270.7964}]}, {"code": "PLY-001-12", "desc": "\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0628\u0644\u0627\u0649", "conv": 12, "sp": 413.0, "rawBoxCost": 209.2587, "boxCost": 282.4992, "profit": 130.5008, "profitPct": 31.6, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PLY-001-12", "qty": 1.0, "unitCost": 7.5, "lineCost": 7.5}, {"child": "SF-PLY-001", "qty": 12.0, "unitCost": 15.5, "lineCost": 186.0}, {"child": "ST-PLY-001", "qty": 12.0, "unitCost": 1.31322314049587, "lineCost": 15.7587}]}, {"code": "PMP-001-24", "desc": "\u0645\u0646\u0638\u0645 \u0645\u0639\u0627\u0644\u0642 \u0628\u064a\u0631\u0641\u0643\u062a ", "conv": 24, "sp": 2250.0, "rawBoxCost": 498.9495, "boxCost": 673.5818, "profit": 1576.4182, "profitPct": 70.06, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PMP-001-24", "qty": 1.0, "unitCost": 34.1030468650236, "lineCost": 34.103}, {"child": "MB-018", "qty": 0.02184, "unitCost": 100.427215189873, "lineCost": 2.1933}, {"child": "MB-030", "qty": 0.02184, "unitCost": 223.736576152305, "lineCost": 4.8864}, {"child": "MB-037", "qty": 0.02184, "unitCost": 109.956331877729, "lineCost": 2.4014}, {"child": "MB-038", "qty": 0.0744, "unitCost": 185.610749185668, "lineCost": 13.8094}, {"child": "MB-040", "qty": 0.02184, "unitCost": 114.00269541779, "lineCost": 2.4898}, {"child": "MB-045", "qty": 0.02448, "unitCost": 190.492957746479, "lineCost": 4.6633}, {"child": "MB-046", "qty": 0.02184, "unitCost": 112.599364069952, "lineCost": 2.4592}, {"child": "MB-047", "qty": 0.02184, "unitCost": 114.0, "lineCost": 2.4898}, {"child": "PPK-101", "qty": 8.04, "unitCost": 51.4602533485619, "lineCost": 413.7404}, {"child": "ST-PMP-001", "qty": 24.0, "unitCost": 0.654729537258705, "lineCost": 15.7135}]}, {"code": "POL-001-00", "desc": "\u0639\u0644\u0628\u0629 \u0628\u0648\u0644\u064a\u0634 \u0643\u0627\u0645\u0644\u0629", "conv": 1, "sp": 0, "rawBoxCost": 2.8833, "boxCost": 3.8925, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "MB-039", "qty": 0.00024, "unitCost": 96.0381355932203, "lineCost": 0.023}, {"child": "MB-210", "qty": 0.0003, "unitCost": 203.333333333333, "lineCost": 0.061}, {"child": "PET-001", "qty": 0.03, "unitCost": 61.9588348579682, "lineCost": 1.8588}, {"child": "PPH-001", "qty": 0.008, "unitCost": 67.4507739938081, "lineCost": 0.5396}, {"child": "PPR-001", "qty": 0.005, "unitCost": 80.1845484377229, "lineCost": 0.4009}]}, {"code": "POP-001-12", "desc": "\u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u0635\u063a\u064a\u0631", "conv": 144, "sp": 1344.0, "rawBoxCost": 378.9211, "boxCost": 511.5435, "profit": 832.4565, "profitPct": 61.94, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-001-144", "qty": 1.0, "unitCost": 17.6820970537262, "lineCost": 17.6821}, {"child": "MB-016", "qty": 0.0288, "unitCost": 118.069651741294, "lineCost": 3.4004}, {"child": "MB-017", "qty": 0.0288, "unitCost": 95.0, "lineCost": 2.736}, {"child": "MB-022", "qty": 0.0288, "unitCost": 130.0, "lineCost": 3.744}, {"child": "MB-027", "qty": 0.0288, "unitCost": 130.0, "lineCost": 3.744}, {"child": "MB-028", "qty": 0.0288, "unitCost": 130.0, "lineCost": 3.744}, {"child": "MB-040", "qty": 0.0288, "unitCost": 114.00269541779, "lineCost": 3.2833}, {"child": "PPH-005", "qty": 5.76, "unitCost": 55.5896038716616, "lineCost": 320.1961}, {"child": "ST-POP-001", "qty": 144.0, "unitCost": 0.141605783866058, "lineCost": 20.3912}]}, {"code": "POP-002-12", "desc": "\u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u0648\u0633\u0637", "conv": 108, "sp": 1550.0, "rawBoxCost": 446.5106, "boxCost": 602.7893, "profit": 947.2107, "profitPct": 61.11, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-002-108", "qty": 1.0, "unitCost": 16.0141905499613, "lineCost": 16.0142}, {"child": "MB-016", "qty": 0.0351, "unitCost": 118.069651741294, "lineCost": 4.1442}, {"child": "MB-022", "qty": 0.0351, "unitCost": 130.0, "lineCost": 4.563}, {"child": "MB-025", "qty": 0.0351, "unitCost": 97.5, "lineCost": 3.4223}, {"child": "MB-027", "qty": 0.0351, "unitCost": 130.0, "lineCost": 4.563}, {"child": "MB-028", "qty": 0.0351, "unitCost": 130.0, "lineCost": 4.563}, {"child": "MB-040", "qty": 0.0351, "unitCost": 114.00269541779, "lineCost": 4.0015}, {"child": "PPH-005", "qty": 7.02, "unitCost": 55.5896038716616, "lineCost": 390.239}, {"child": "ST-POP-002", "qty": 108.0, "unitCost": 0.138892857142857, "lineCost": 15.0004}]}, {"code": "POP-003-12", "desc": "\u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u0643\u0628\u064a\u0631", "conv": 48, "sp": 1125.0, "rawBoxCost": 332.5142, "boxCost": 448.8942, "profit": 676.1058, "profitPct": 60.1, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-003-48", "qty": 1.0, "unitCost": 16.24, "lineCost": 16.24}, {"child": "MB-016", "qty": 0.0264, "unitCost": 118.069651741294, "lineCost": 3.117}, {"child": "MB-017", "qty": 0.0264, "unitCost": 95.0, "lineCost": 2.508}, {"child": "MB-027", "qty": 0.0264, "unitCost": 130.0, "lineCost": 3.432}, {"child": "MB-028", "qty": 0.0264, "unitCost": 130.0, "lineCost": 3.432}, {"child": "MB-040", "qty": 0.024, "unitCost": 114.00269541779, "lineCost": 2.7361}, {"child": "MB-050", "qty": 0.024, "unitCost": 114.0, "lineCost": 2.736}, {"child": "PPH-005", "qty": 5.28, "unitCost": 55.5896038716616, "lineCost": 293.5131}, {"child": "ST-POP-103", "qty": 48.0, "unitCost": 0.1, "lineCost": 4.8}]}, {"code": "POP-101-36", "desc": "\u0637\u0642\u0645 \u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u062b\u0644\u0627\u062b\u064a \u0635\u063a\u064a\u0631 \u0639\u062f\u062f 36 \u0637\u0642\u0645", "conv": 36, "sp": 0, "rawBoxCost": 276.918, "boxCost": 373.8393, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-001-144", "qty": 1.0, "unitCost": 17.6820970537262, "lineCost": 17.6821}, {"child": "MB-016", "qty": 0.0216, "unitCost": 118.069651741294, "lineCost": 2.5503}, {"child": "MB-017", "qty": 0.0216, "unitCost": 95.0, "lineCost": 2.052}, {"child": "MB-022", "qty": 0.0216, "unitCost": 130.0, "lineCost": 2.808}, {"child": "MB-027", "qty": 0.0216, "unitCost": 130.0, "lineCost": 2.808}, {"child": "MB-028", "qty": 0.0216, "unitCost": 130.0, "lineCost": 2.808}, {"child": "MB-040", "qty": 0.0216, "unitCost": 114.00269541779, "lineCost": 2.4625}, {"child": "PPH-005", "qty": 4.32, "unitCost": 55.5896038716616, "lineCost": 240.1471}, {"child": "ST-POP-101", "qty": 36.0, "unitCost": 0.1, "lineCost": 3.6}]}, {"code": "POP-102-30", "desc": "\u0637\u0642\u0645 \u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u062b\u0644\u0627\u062b\u064a \u0648\u0633\u0637 \u0639\u062f\u062f 30 \u0637\u0642\u0645 ", "conv": 30, "sp": 0, "rawBoxCost": 365.2609, "boxCost": 493.1022, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-002-108", "qty": 1.0, "unitCost": 16.0141905499613, "lineCost": 16.0142}, {"child": "MB-016", "qty": 0.02925, "unitCost": 118.069651741294, "lineCost": 3.4535}, {"child": "MB-022", "qty": 0.02925, "unitCost": 130.0, "lineCost": 3.8025}, {"child": "MB-025", "qty": 0.02925, "unitCost": 97.5, "lineCost": 2.8519}, {"child": "MB-027", "qty": 0.02925, "unitCost": 130.0, "lineCost": 3.8025}, {"child": "MB-028", "qty": 0.02925, "unitCost": 130.0, "lineCost": 3.8025}, {"child": "MB-040", "qty": 0.02925, "unitCost": 114.00269541779, "lineCost": 3.3346}, {"child": "PPH-005", "qty": 5.85, "unitCost": 55.5896038716616, "lineCost": 325.1992}, {"child": "ST-POP-102", "qty": 30.0, "unitCost": 0.1, "lineCost": 3.0}]}, {"code": "POP-103- 12", "desc": "\u0637\u0642\u0645 \u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u062b\u0644\u0627\u062b\u064a \u0643\u0628\u064a\u0631 \u0639\u062f\u062f 12 \u0637\u0642\u0645 ", "conv": 12, "sp": 0, "rawBoxCost": 251.5676, "boxCost": 339.6163, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-003-48", "qty": 1.0, "unitCost": 16.24, "lineCost": 16.24}, {"child": "MB-016", "qty": 0.0198, "unitCost": 118.069651741294, "lineCost": 2.3378}, {"child": "MB-017", "qty": 0.0198, "unitCost": 95.0, "lineCost": 1.881}, {"child": "MB-022", "qty": 0.0198, "unitCost": 130.0, "lineCost": 2.574}, {"child": "MB-027", "qty": 0.0198, "unitCost": 130.0, "lineCost": 2.574}, {"child": "MB-028", "qty": 0.0198, "unitCost": 130.0, "lineCost": 2.574}, {"child": "MB-040", "qty": 0.018, "unitCost": 114.00269541779, "lineCost": 2.052}, {"child": "PPH-005", "qty": 3.96, "unitCost": 55.5896038716616, "lineCost": 220.1348}, {"child": "ST-POP-103", "qty": 12.0, "unitCost": 0.1, "lineCost": 1.2}]}, {"code": "POP-104-12", "desc": "\u0637\u0642\u0645 \u0637\u0628\u0642 \u0627\u0648\u0641\u0627\u0644 \u062b\u0644\u0627\u062b\u064a \u0645\u0634\u0643\u0644", "conv": 24, "sp": 1100.0, "rawBoxCost": 333.8226, "boxCost": 450.6605, "profit": 649.3395, "profitPct": 59.03, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-POP-101-24", "qty": 1.0, "unitCost": 14.5, "lineCost": 14.5}, {"child": "MB-016", "qty": 0.0258, "unitCost": 118.069651741294, "lineCost": 3.0462}, {"child": "MB-017", "qty": 0.018, "unitCost": 95.0, "lineCost": 1.71}, {"child": "MB-022", "qty": 0.0258, "unitCost": 130.0, "lineCost": 3.354}, {"child": "MB-025", "qty": 0.0078, "unitCost": 97.5, "lineCost": 0.7605}, {"child": "MB-028", "qty": 0.0258, "unitCost": 130.0, "lineCost": 3.354}, {"child": "MB-040", "qty": 0.0246, "unitCost": 114.00269541779, "lineCost": 2.8045}, {"child": "MB-050", "qty": 0.0168, "unitCost": 114.0, "lineCost": 1.9152}, {"child": "MB-051", "qty": 0.0078, "unitCost": 111.527777777778, "lineCost": 0.8699}, {"child": "PPH-005", "qty": 5.16, "unitCost": 55.5896038716616, "lineCost": 286.8424}, {"child": "ST-POP-104", "qty": 24.0, "unitCost": 0.611081081081081, "lineCost": 14.6659}]}, {"code": "PRM-001-06", "desc": "\u0639\u0644\u0628\u0629 \u062a\u0648\u0627\u0628\u0644 150 \u0645\u0644", "conv": 6, "sp": 2138.0, "rawBoxCost": 831.4642, "boxCost": 1122.4767, "profit": 1015.5233, "profitPct": 47.5, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PRM-901-99", "qty": 1.0, "unitCost": 22.4115867158672, "lineCost": 22.4116}, {"child": "CI-PRM-901-99", "qty": 6.0, "unitCost": 10.3234162365087, "lineCost": 61.9405}, {"child": "MB-026", "qty": 0.0144, "unitCost": 316.0, "lineCost": 4.5504}, {"child": "MB-030", "qty": 0.01584, "unitCost": 223.736576152305, "lineCost": 3.544}, {"child": "MB-037", "qty": 0.01584, "unitCost": 109.956331877729, "lineCost": 1.7417}, {"child": "MB-038", "qty": 0.0185, "unitCost": 185.610749185668, "lineCost": 3.4338}, {"child": "MB-039", "qty": 0.00615, "unitCost": 96.0381355932203, "lineCost": 0.5906}, {"child": "MB-040", "qty": 0.02484, "unitCost": 114.00269541779, "lineCost": 2.8318}, {"child": "MB-105", "qty": 0.0144, "unitCost": 265.0, "lineCost": 3.816}, {"child": "PPH-001", "qty": 4.12056, "unitCost": 67.4507739938081, "lineCost": 277.935}, {"child": "PPR-001", "qty": 4.752, "unitCost": 80.1845484377229, "lineCost": 381.037}, {"child": "ST-PRM-001", "qty": 216.0, "unitCost": 0.313110236220472, "lineCost": 67.6318}]}, {"code": "PRM-002-06", "desc": "\u0639\u0644\u0628\u0629 \u062a\u0648\u0627\u0628\u0644 300 \u0645\u0644", "conv": 6, "sp": 1950.0, "rawBoxCost": 694.8644, "boxCost": 938.0669, "profit": 1011.9331, "profitPct": 51.89, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PRM-901-99", "qty": 1.0, "unitCost": 22.4115867158672, "lineCost": 22.4116}, {"child": "CI-PRM-901-99", "qty": 6.0, "unitCost": 10.3234162365087, "lineCost": 61.9405}, {"child": "MB-026", "qty": 0.0096, "unitCost": 316.0, "lineCost": 3.0336}, {"child": "MB-030", "qty": 0.01056, "unitCost": 223.736576152305, "lineCost": 2.3627}, {"child": "MB-037", "qty": 0.01056, "unitCost": 109.956331877729, "lineCost": 1.1611}, {"child": "MB-038", "qty": 0.01234, "unitCost": 185.610749185668, "lineCost": 2.2904}, {"child": "MB-039", "qty": 0.00411, "unitCost": 96.0381355932203, "lineCost": 0.3947}, {"child": "MB-040", "qty": 0.01656, "unitCost": 114.00269541779, "lineCost": 1.8879}, {"child": "MB-105", "qty": 0.0096, "unitCost": 265.0, "lineCost": 2.544}, {"child": "PPH-001", "qty": 2.74704, "unitCost": 67.4507739938081, "lineCost": 185.29}, {"child": "PPR-001", "qty": 4.32, "unitCost": 80.1845484377229, "lineCost": 346.3972}, {"child": "ST-PRM-002", "qty": 144.0, "unitCost": 0.452435233160622, "lineCost": 65.1507}]}, {"code": "PRM-101-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u062a\u0648\u0627\u0628\u0644 \u0628\u0631\u0627\u064a\u0645 \u0631\u0628\u0627\u0639\u064a 150 \u0645\u0644", "conv": 12, "sp": 863.0, "rawBoxCost": 298.1669, "boxCost": 402.5253, "profit": 460.4747, "profitPct": 53.36, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PRM-101-12", "qty": 1.0, "unitCost": 19.7802770083103, "lineCost": 19.7803}, {"child": "MB-026", "qty": 0.0101, "unitCost": 316.0, "lineCost": 3.1916}, {"child": "MB-030", "qty": 0.01042, "unitCost": 223.736576152305, "lineCost": 2.3313}, {"child": "MB-037", "qty": 0.01042, "unitCost": 109.956331877729, "lineCost": 1.1457}, {"child": "MB-038", "qty": 0.01102, "unitCost": 185.610749185668, "lineCost": 2.0454}, {"child": "MB-039", "qty": 0.00138, "unitCost": 96.0381355932203, "lineCost": 0.1325}, {"child": "MB-040", "qty": 0.01242, "unitCost": 114.00269541779, "lineCost": 1.4159}, {"child": "MB-105", "qty": 0.0101, "unitCost": 265.0, "lineCost": 2.6765}, {"child": "PPH-001", "qty": 0.91568, "unitCost": 67.4507739938081, "lineCost": 61.7633}, {"child": "PPK-001", "qty": 1.38, "unitCost": 75.3691770186335, "lineCost": 104.0095}, {"child": "PPR-001", "qty": 1.056, "unitCost": 80.1845484377229, "lineCost": 84.6749}, {"child": "ST-PRM-101", "qty": 12.0, "unitCost": 1.25, "lineCost": 15.0}]}, {"code": "PRM-102-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u062a\u0648\u0627\u0628\u0644 \u0628\u0631\u0627\u064a\u0645 \u0631\u0628\u0627\u0639\u064a 300 \u0645\u0644", "conv": 12, "sp": 1038.0, "rawBoxCost": 19.44, "boxCost": 26.244, "profit": 1011.756, "profitPct": 97.47, "hasBOM": true, "missing": ["CI-PRM-102-01", "SU-PRM-102-01", "SU-PRM-102-02", "SU-PRM-102-03", "SU-PRM-102-04", "SU-PRM-102-05", "SU-PRM-102-06"], "bomLines": [{"child": "CE-PRM-102-12", "qty": 1.0, "unitCost": 19.44, "lineCost": 19.44}, {"child": "CI-PRM-102-01", "qty": 12.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-01", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-02", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-03", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-04", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-05", "qty": 2.0, "unitCost": null, "lineCost": 0.0}, {"child": "SU-PRM-102-06", "qty": 2.0, "unitCost": null, "lineCost": 0.0}]}, {"code": "PRM-103-12", "desc": "\u0633\u062a\u0627\u0646\u062f \u062a\u0648\u0627\u0628\u0644 \u0628\u0631\u0627\u064a\u0645 \u062b\u0645\u0627\u0646\u064a \u0645\u0634\u0643\u0644", "conv": 12, "sp": 1706.0, "rawBoxCost": 581.7153, "boxCost": 785.3157, "profit": 920.6843, "profitPct": 53.97, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PRM-103-12", "qty": 1.0, "unitCost": 24.9806358974359, "lineCost": 24.9806}, {"child": "MB-014", "qty": 0.0132, "unitCost": 105.0, "lineCost": 1.386}, {"child": "MB-026", "qty": 0.0064, "unitCost": 316.0, "lineCost": 2.0224}, {"child": "MB-030", "qty": 0.02024, "unitCost": 223.736576152305, "lineCost": 4.5284}, {"child": "MB-037", "qty": 0.02024, "unitCost": 109.956331877729, "lineCost": 2.2255}, {"child": "MB-038", "qty": 0.01812, "unitCost": 185.610749185668, "lineCost": 3.3633}, {"child": "MB-039", "qty": 0.00603, "unitCost": 96.0381355932203, "lineCost": 0.5791}, {"child": "MB-040", "qty": 0.02424, "unitCost": 114.00269541779, "lineCost": 2.7634}, {"child": "MB-105", "qty": 0.0196, "unitCost": 265.0, "lineCost": 5.194}, {"child": "PPH-001", "qty": 1.83136, "unitCost": 67.4507739938081, "lineCost": 123.5266}, {"child": "PPK-001", "qty": 2.5608, "unitCost": 75.3691770186335, "lineCost": 193.0054}, {"child": "PPR-001", "qty": 2.496, "unitCost": 80.1845484377229, "lineCost": 200.1406}, {"child": "ST-PRM-103", "qty": 12.0, "unitCost": 1.5, "lineCost": 18.0}]}, {"code": "PSH-001-01", "desc": "\u062c\u0632\u0627\u0645\u0629 5 \u062f\u0648\u0631 \u0628\u064a\u0631\u0641\u0643\u062a", "conv": 1, "sp": 275.0, "rawBoxCost": 104.953, "boxCost": 141.6866, "profit": 133.3134, "profitPct": 48.48, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PSH-001-01", "qty": 1.0, "unitCost": 15.5627264492754, "lineCost": 15.5627}, {"child": "MB-030", "qty": 0.00024, "unitCost": 223.736576152305, "lineCost": 0.0537}, {"child": "MB-037", "qty": 0.045, "unitCost": 109.956331877729, "lineCost": 4.948}, {"child": "PPH-004", "qty": 1.92, "unitCost": 43.7207771717081, "lineCost": 83.9439}, {"child": "PPH-005", "qty": 0.008, "unitCost": 55.5896038716616, "lineCost": 0.4447}]}, {"code": "psh-002-01", "desc": "\u062c\u0632\u0627\u0645\u0629 4\u062f\u0648\u0631 \u0628\u064a\u0631\u0641\u0643\u062a", "conv": 1, "sp": 0, "rawBoxCost": 102.0047, "boxCost": 137.7063, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PSH-002-01", "qty": 1.0, "unitCost": 16.5, "lineCost": 16.5}, {"child": "MB-030", "qty": 0.00024, "unitCost": 223.736576152305, "lineCost": 0.0537}, {"child": "MB-037", "qty": 0.036, "unitCost": 109.956331877729, "lineCost": 3.9584}, {"child": "PPH-004", "qty": 0.328, "unitCost": 43.7207771717081, "lineCost": 14.3404}, {"child": "PPH-005", "qty": 1.208, "unitCost": 55.5896038716616, "lineCost": 67.1522}]}, {"code": "PSH-004-01", "desc": "\u062c\u0632\u0627\u0645\u0629 6 \u062f\u0648\u0631 \u0628\u064a\u0631\u0641\u0643\u062a", "conv": 1, "sp": 0, "rawBoxCost": 147.3823, "boxCost": 198.9661, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PSH-003-01", "qty": 1.0, "unitCost": 18.5, "lineCost": 18.5}, {"child": "MB-030", "qty": 0.00024, "unitCost": 223.736576152305, "lineCost": 0.0537}, {"child": "MB-037", "qty": 0.054, "unitCost": 109.956331877729, "lineCost": 5.9376}, {"child": "PPH-004", "qty": 0.512, "unitCost": 43.7207771717081, "lineCost": 22.385}, {"child": "PPH-005", "qty": 1.808, "unitCost": 55.5896038716616, "lineCost": 100.506}]}, {"code": "PSP-001-48", "desc": "\u0637\u0628\u0642 \u0633\u0631\u0641\u064a\u0633", "conv": 48, "sp": 1219.0, "rawBoxCost": 429.2732, "boxCost": 579.5188, "profit": 639.4812, "profitPct": 52.46, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PAP-001-60", "qty": 1.0, "unitCost": 9.4695514092894, "lineCost": 9.4696}, {"child": "MB-016", "qty": 0.036, "unitCost": 118.069651741294, "lineCost": 4.2505}, {"child": "MB-021", "qty": 0.036, "unitCost": 126.344086021505, "lineCost": 4.5484}, {"child": "MB-028", "qty": 0.036, "unitCost": 130.0, "lineCost": 4.68}, {"child": "MB-040", "qty": 0.036, "unitCost": 114.00269541779, "lineCost": 4.1041}, {"child": "MB-045", "qty": 0.036, "unitCost": 190.492957746479, "lineCost": 6.8577}, {"child": "MB-050", "qty": 0.036, "unitCost": 114.0, "lineCost": 4.104}, {"child": "PPH-005", "qty": 6.952, "unitCost": 55.5896038716616, "lineCost": 386.4589}, {"child": "ST-PSP-001-1", "qty": 48.0, "unitCost": 0.1, "lineCost": 4.8}]}, {"code": "PST-001-12", "desc": "\u0645\u0635\u0641\u0649 2 \u0641\u064a 1", "conv": 12, "sp": 813.0, "rawBoxCost": 270.6515, "boxCost": 365.3795, "profit": 447.6205, "profitPct": 55.06, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PST-001-12", "qty": 1.0, "unitCost": 12.8474499604207, "lineCost": 12.8474}, {"child": "MB-018", "qty": 0.012, "unitCost": 100.427215189873, "lineCost": 1.2051}, {"child": "MB-023", "qty": 0.012, "unitCost": 130.0, "lineCost": 1.56}, {"child": "MB-026", "qty": 0.0093, "unitCost": 316.0, "lineCost": 2.9388}, {"child": "MB-038", "qty": 0.0348, "unitCost": 185.610749185668, "lineCost": 6.4593}, {"child": "MB-039", "qty": 0.0117, "unitCost": 96.0381355932203, "lineCost": 1.1236}, {"child": "MB-040", "qty": 0.012, "unitCost": 114.00269541779, "lineCost": 1.368}, {"child": "MB-046", "qty": 0.012, "unitCost": 112.599364069952, "lineCost": 1.3512}, {"child": "MB-048", "qty": 0.006, "unitCost": 110.0, "lineCost": 0.66}, {"child": "MB-049", "qty": 0.006, "unitCost": 110.0, "lineCost": 0.66}, {"child": "MB-050", "qty": 0.006, "unitCost": 114.0, "lineCost": 0.684}, {"child": "MB-051", "qty": 0.006, "unitCost": 111.527777777778, "lineCost": 0.6692}, {"child": "PPH-005", "qty": 4.2042, "unitCost": 55.5896038716616, "lineCost": 233.7098}, {"child": "ST-PST-001", "qty": 12.0, "unitCost": 0.451258470474347, "lineCost": 5.4151}]}, {"code": "PTR-001-06", "desc": "\u062a\u0631\u0648\u0644\u064a \u0628\u064a\u0631\u0641\u0643\u062a 4 \u062f\u0648\u0631", "conv": 6, "sp": 1625.0, "rawBoxCost": 498.7565, "boxCost": 673.3213, "profit": 951.6787, "profitPct": 58.56, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PTR", "qty": 1.0, "unitCost": 27.2584547346514, "lineCost": 27.2585}, {"child": "MB-013", "qty": 0.04624, "unitCost": 190.306387225549, "lineCost": 8.7998}, {"child": "MB-015", "qty": 0.03312, "unitCost": 119.240506329114, "lineCost": 3.9492}, {"child": "MB-018", "qty": 0.0068, "unitCost": 100.427215189873, "lineCost": 0.6829}, {"child": "MB-020", "qty": 0.01336, "unitCost": 121.964285714286, "lineCost": 1.6294}, {"child": "MB-030", "qty": 0.0336, "unitCost": 223.736576152305, "lineCost": 7.5175}, {"child": "MB-037", "qty": 0.03992, "unitCost": 109.956331877729, "lineCost": 4.3895}, {"child": "MB-038", "qty": 0.00024, "unitCost": 185.610749185668, "lineCost": 0.0445}, {"child": "MB-040", "qty": 0.03312, "unitCost": 114.00269541779, "lineCost": 3.7758}, {"child": "MB-045", "qty": 0.03312, "unitCost": 190.492957746479, "lineCost": 6.3091}, {"child": "PPH-003", "qty": 0.27048, "unitCost": 32.0, "lineCost": 8.6554}, {"child": "PPH-004", "qty": 0.656, "unitCost": 43.7207771717081, "lineCost": 28.6808}, {"child": "PPH-005", "qty": 0.72872, "unitCost": 55.5896038716616, "lineCost": 40.5093}, {"child": "PPH-007", "qty": 0.32144, "unitCost": 35.1771189071717, "lineCost": 11.3073}, {"child": "PPK-101", "qty": 6.624, "unitCost": 51.4602533485619, "lineCost": 340.8727}, {"child": "ST-PTR-001", "qty": 6.0, "unitCost": 0.729127125810392, "lineCost": 4.3748}]}, {"code": "PTR-002-06", "desc": "\u062a\u0631\u0648\u0644\u064a \u0628\u064a\u0631\u0641\u0643\u062a 3\u062f\u0648\u0631", "conv": 6, "sp": 1219.0, "rawBoxCost": 377.4471, "boxCost": 509.5536, "profit": 709.4464, "profitPct": 58.2, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-PTR-002-06", "qty": 1.0, "unitCost": 24.6463556851312, "lineCost": 24.6464}, {"child": "MB-013", "qty": 0.03428, "unitCost": 190.306387225549, "lineCost": 6.5237}, {"child": "MB-015", "qty": 0.02484, "unitCost": 119.240506329114, "lineCost": 2.9619}, {"child": "MB-018", "qty": 0.00496, "unitCost": 100.427215189873, "lineCost": 0.4981}, {"child": "MB-020", "qty": 0.00968, "unitCost": 121.964285714286, "lineCost": 1.1806}, {"child": "MB-030", "qty": 0.02532, "unitCost": 223.736576152305, "lineCost": 5.665}, {"child": "MB-037", "qty": 0.0298, "unitCost": 109.956331877729, "lineCost": 3.2767}, {"child": "MB-038", "qty": 0.00024, "unitCost": 185.610749185668, "lineCost": 0.0445}, {"child": "MB-040", "qty": 0.02484, "unitCost": 114.00269541779, "lineCost": 2.8318}, {"child": "MB-045", "qty": 0.02484, "unitCost": 190.492957746479, "lineCost": 4.7318}, {"child": "PPH-003", "qty": 0.18032, "unitCost": 32.0, "lineCost": 5.7702}, {"child": "PPH-004", "qty": 0.472, "unitCost": 43.7207771717081, "lineCost": 20.6362}, {"child": "PPH-005", "qty": 0.55208, "unitCost": 55.5896038716616, "lineCost": 30.6899}, {"child": "PPH-007", "qty": 0.23128, "unitCost": 35.1771189071717, "lineCost": 8.1358}, {"child": "PPK-101", "qty": 4.968, "unitCost": 51.4602533485619, "lineCost": 255.6545}, {"child": "ST-PTR-002", "qty": 6.0, "unitCost": 0.7, "lineCost": 4.2}]}, {"code": "SHV-001-1500", "desc": "\u0645\u0634\u062d\u0627\u0641 \u062c\u0648\u0644\u062f\u064a\u0646 \u0644\u0648\u0646 \u0627\u062d\u0645\u0631", "conv": 1500, "sp": 0, "rawBoxCost": 1393.6551, "boxCost": 1881.4344, "profit": 0, "profitPct": 0, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SGT-001", "qty": 1.0, "unitCost": 23.5291902071563, "lineCost": 23.5292}, {"child": "PPK-103", "qty": 28.5, "unitCost": 48.0745920745921, "lineCost": 1370.1259}]}, {"code": "SMT-001-12", "desc": "\u062d\u0627\u0641\u0638\u0629 \u0637\u0639\u0627\u0645 \u0633\u0645\u0627\u0631\u062a \u0641\u0631\u062f\u0627\u0646\u064a", "conv": 12, "sp": 858.0, "rawBoxCost": 308.7032, "boxCost": 416.7493, "profit": 441.2507, "profitPct": 51.43, "hasBOM": true, "missing": [], "bomLines": [{"child": "MB-045", "qty": 0.00678, "unitCost": 190.492957746479, "lineCost": 1.2915}, {"child": "PPK-001", "qty": 1.2, "unitCost": 75.3691770186335, "lineCost": 90.443}, {"child": "PPR-001", "qty": 1.74, "unitCost": 80.1845484377229, "lineCost": 139.5211}, {"child": "RUB-001", "qty": 0.312, "unitCost": 182.875, "lineCost": 57.057}, {"child": "ST-SMT-001", "qty": 12.0, "unitCost": 0.333082553883693, "lineCost": 3.997}, {"child": "CE-SMT-001-12", "qty": 1.0, "unitCost": 11.8783429672447, "lineCost": 11.8783}, {"child": "MB-016", "qty": 0.00678, "unitCost": 118.069651741294, "lineCost": 0.8005}, {"child": "MB-020", "qty": 0.00678, "unitCost": 121.964285714286, "lineCost": 0.8269}, {"child": "MB-021", "qty": 0.00678, "unitCost": 126.344086021505, "lineCost": 0.8566}, {"child": "MB-038", "qty": 0.00678, "unitCost": 185.610749185668, "lineCost": 1.2584}, {"child": "MB-040", "qty": 0.00678, "unitCost": 114.00269541779, "lineCost": 0.7729}]}, {"code": "SMT-002-12", "desc": "\u062d\u0627\u0641\u0638\u0629 \u0637\u0639\u0627\u0645 \u0633\u0645\u0627\u0631\u062a \u062b\u0646\u0627\u0626\u064a", "conv": 12, "sp": 1413.0, "rawBoxCost": 491.7784, "boxCost": 663.9008, "profit": 749.0992, "profitPct": 53.01, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SMT-002-12", "qty": 1.0, "unitCost": 16.5775862068966, "lineCost": 16.5776}, {"child": "MB-016", "qty": 0.00816, "unitCost": 118.069651741294, "lineCost": 0.9634}, {"child": "MB-020", "qty": 0.00816, "unitCost": 121.964285714286, "lineCost": 0.9952}, {"child": "MB-021", "qty": 0.00816, "unitCost": 126.344086021505, "lineCost": 1.031}, {"child": "MB-038", "qty": 0.00816, "unitCost": 185.610749185668, "lineCost": 1.5146}, {"child": "MB-040", "qty": 0.00816, "unitCost": 114.00269541779, "lineCost": 0.9303}, {"child": "MB-045", "qty": 0.00816, "unitCost": 190.492957746479, "lineCost": 1.5544}, {"child": "PPK-001", "qty": 1.32, "unitCost": 75.3691770186335, "lineCost": 99.4873}, {"child": "PPR-001", "qty": 3.48, "unitCost": 80.1845484377229, "lineCost": 279.0422}, {"child": "RUB-001", "qty": 0.468, "unitCost": 182.875, "lineCost": 85.5855}, {"child": "ST-SMT-002", "qty": 12.0, "unitCost": 0.341411089288412, "lineCost": 4.0969}]}, {"code": "SMT-003-12", "desc": "\u062d\u0627\u0641\u0638\u0629 \u0637\u0639\u0627\u0645 \u0633\u0645\u0627\u0631\u062a \u062b\u0644\u0627\u062b\u064a", "conv": 12, "sp": 1875.0, "rawBoxCost": 669.7797, "boxCost": 904.2026, "profit": 970.7974, "profitPct": 51.78, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SMT-003-12", "qty": 1.0, "unitCost": 15.0, "lineCost": 15.0}, {"child": "MB-016", "qty": 0.00954, "unitCost": 118.069651741294, "lineCost": 1.1264}, {"child": "MB-020", "qty": 0.00954, "unitCost": 121.964285714286, "lineCost": 1.1635}, {"child": "MB-021", "qty": 0.00954, "unitCost": 126.344086021505, "lineCost": 1.2053}, {"child": "MB-038", "qty": 0.00954, "unitCost": 185.610749185668, "lineCost": 1.7707}, {"child": "MB-040", "qty": 0.00954, "unitCost": 114.00269541779, "lineCost": 1.0876}, {"child": "MB-045", "qty": 0.00954, "unitCost": 190.492957746479, "lineCost": 1.8173}, {"child": "PPK-001", "qty": 1.44, "unitCost": 75.3691770186335, "lineCost": 108.5316}, {"child": "PPR-001", "qty": 5.22, "unitCost": 80.1845484377229, "lineCost": 418.5633}, {"child": "RUB-001", "qty": 0.624, "unitCost": 182.875, "lineCost": 114.114}, {"child": "ST-SMT-003", "qty": 12.0, "unitCost": 0.45, "lineCost": 5.4}]}, {"code": "SQU-001-24", "desc": "\u0645\u062c\u0645\u0648\u0639\u0647 \u0625\u0632\u0627\u0632\u0629 \u0645\u064a\u0627\u0629 \u0648\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0647\u064a\u0631\u0648", "conv": 24, "sp": 1500.0, "rawBoxCost": 612.2768, "boxCost": 826.5737, "profit": 673.4263, "profitPct": 44.9, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SQU-001-24", "qty": 1.0, "unitCost": 17.0283035714286, "lineCost": 17.0283}, {"child": "M50-001", "qty": 0.12, "unitCost": 74.4253393665158, "lineCost": 8.931}, {"child": "MB-032", "qty": 0.09384, "unitCost": 190.363636363636, "lineCost": 17.8637}, {"child": "MB-035", "qty": 0.09384, "unitCost": 186.397467572576, "lineCost": 17.4915}, {"child": "MB-036", "qty": 0.09384, "unitCost": 188.125, "lineCost": 17.6537}, {"child": "MB-039", "qty": 0.0432, "unitCost": 96.0381355932203, "lineCost": 4.1488}, {"child": "MB-043", "qty": 0.09384, "unitCost": 170.799180327869, "lineCost": 16.0278}, {"child": "MB-044", "qty": 0.09384, "unitCost": 188.547619047619, "lineCost": 17.6933}, {"child": "MB-055", "qty": 0.09384, "unitCost": 214.316753926702, "lineCost": 20.1115}, {"child": "MB-201", "qty": 0.0016, "unitCost": 323.888888888889, "lineCost": 0.5182}, {"child": "MB-202", "qty": 0.0016, "unitCost": 353.333333333333, "lineCost": 0.5653}, {"child": "MB-203", "qty": 0.0016, "unitCost": 344.815950920245, "lineCost": 0.5517}, {"child": "MB-205", "qty": 0.0016, "unitCost": 297.352941176471, "lineCost": 0.4758}, {"child": "MB-206", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "MB-207", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "PET-001", "qty": 1.2, "unitCost": 61.9588348579682, "lineCost": 74.3506}, {"child": "PPH-001", "qty": 1.752, "unitCost": 67.4507739938081, "lineCost": 118.1738}, {"child": "PPH-005", "qty": 0.66, "unitCost": 55.5896038716616, "lineCost": 36.6891}, {"child": "PPR-001", "qty": 2.76, "unitCost": 80.1845484377229, "lineCost": 221.3094}, {"child": "ST-PLN-001", "qty": 24.0, "unitCost": 0.890222222222222, "lineCost": 21.3653}]}, {"code": "SQU-002-24", "desc": "\u0645\u062c\u0645\u0648\u0639\u0647 \u0625\u0632\u0627\u0632\u0629 \u0645\u064a\u0627\u0629 \u0648\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0643\u064a\u062f\u0632", "conv": 24, "sp": 930.0, "rawBoxCost": 399.2603, "boxCost": 539.0014, "profit": 390.9986, "profitPct": 42.04, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SQU-002-24", "qty": 1.0, "unitCost": 15.2508143322476, "lineCost": 15.2508}, {"child": "M50-001", "qty": 0.12, "unitCost": 74.4253393665158, "lineCost": 8.931}, {"child": "MB-032", "qty": 0.09204, "unitCost": 190.363636363636, "lineCost": 17.5211}, {"child": "MB-035", "qty": 0.09204, "unitCost": 186.397467572576, "lineCost": 17.156}, {"child": "MB-036", "qty": 0.09204, "unitCost": 188.125, "lineCost": 17.315}, {"child": "MB-043", "qty": 0.09204, "unitCost": 170.799180327869, "lineCost": 15.7204}, {"child": "MB-044", "qty": 0.09204, "unitCost": 188.547619047619, "lineCost": 17.3539}, {"child": "MB-055", "qty": 0.09204, "unitCost": 214.316753926702, "lineCost": 19.7257}, {"child": "MB-201", "qty": 0.0016, "unitCost": 323.888888888889, "lineCost": 0.5182}, {"child": "MB-202", "qty": 0.0016, "unitCost": 353.333333333333, "lineCost": 0.5653}, {"child": "MB-203", "qty": 0.0016, "unitCost": 344.815950920245, "lineCost": 0.5517}, {"child": "MB-205", "qty": 0.0016, "unitCost": 297.352941176471, "lineCost": 0.4758}, {"child": "MB-206", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "MB-207", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "PET-001", "qty": 1.2, "unitCost": 61.9588348579682, "lineCost": 74.3506}, {"child": "PPH-001", "qty": 0.312, "unitCost": 67.4507739938081, "lineCost": 21.0446}, {"child": "PPR-001", "qty": 1.8, "unitCost": 80.1845484377229, "lineCost": 144.3322}, {"child": "ST-AQU-002", "qty": 24.0, "unitCost": 0.48, "lineCost": 11.52}, {"child": "ST-PLN-002", "qty": 24.0, "unitCost": 0.65, "lineCost": 15.6}]}, {"code": "SQU-003-24", "desc": "\u0645\u062c\u0645\u0648\u0639\u0647 \u0625\u0632\u0627\u0632\u0629 \u0645\u064a\u0627\u0629 \u0648\u0644\u0627\u0646\u0634 \u0628\u0648\u0643\u0633 \u0628\u0644\u0627\u0649", "conv": 24, "sp": 1241.0, "rawBoxCost": 544.6221, "boxCost": 735.2398, "profit": 505.7602, "profitPct": 40.75, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SQU-003-24", "qty": 1.0, "unitCost": 19.4969198088157, "lineCost": 19.4969}, {"child": "M50-001", "qty": 0.12, "unitCost": 74.4253393665158, "lineCost": 8.931}, {"child": "MB-032", "qty": 0.00204, "unitCost": 190.363636363636, "lineCost": 0.3883}, {"child": "MB-035", "qty": 0.00204, "unitCost": 186.397467572576, "lineCost": 0.3803}, {"child": "MB-036", "qty": 0.00204, "unitCost": 188.125, "lineCost": 0.3838}, {"child": "MB-043", "qty": 0.00204, "unitCost": 170.799180327869, "lineCost": 0.3484}, {"child": "MB-044", "qty": 0.00204, "unitCost": 188.547619047619, "lineCost": 0.3846}, {"child": "MB-055", "qty": 0.00204, "unitCost": 214.316753926702, "lineCost": 0.4372}, {"child": "MB-201", "qty": 0.0016, "unitCost": 323.888888888889, "lineCost": 0.5182}, {"child": "MB-202", "qty": 0.0016, "unitCost": 353.333333333333, "lineCost": 0.5653}, {"child": "MB-203", "qty": 0.0016, "unitCost": 344.815950920245, "lineCost": 0.5517}, {"child": "MB-205", "qty": 0.0016, "unitCost": 297.352941176471, "lineCost": 0.4758}, {"child": "MB-206", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "MB-207", "qty": 0.0016, "unitCost": 415.0, "lineCost": 0.664}, {"child": "PET-001", "qty": 1.2, "unitCost": 61.9588348579682, "lineCost": 74.3506}, {"child": "PPH-001", "qty": 0.312, "unitCost": 67.4507739938081, "lineCost": 21.0446}, {"child": "SF-PLY-001", "qty": 24.0, "unitCost": 15.5, "lineCost": 372.0}, {"child": "ST-AQU-002", "qty": 24.0, "unitCost": 0.48, "lineCost": 11.52}, {"child": "ST-PLY-001", "qty": 24.0, "unitCost": 1.31322314049587, "lineCost": 31.5174}]}, {"code": "SWT-001-24", "desc": "\u0642\u0627\u0644\u0628 \u062d\u0644\u0648\u064a\u0627\u062a \u0635\u063a\u064a\u0631", "conv": 24, "sp": 1350.0, "rawBoxCost": 407.7359, "boxCost": 550.4435, "profit": 799.5565, "profitPct": 59.23, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SWT-001-24", "qty": 1.0, "unitCost": 10.1525724423418, "lineCost": 10.1526}, {"child": "CI-SWT-001-6", "qty": 24.0, "unitCost": 2.70258828928934, "lineCost": 64.8621}, {"child": "M50-001", "qty": 2.0952, "unitCost": 74.4253393665158, "lineCost": 155.936}, {"child": "MB-021", "qty": 0.0108, "unitCost": 126.344086021505, "lineCost": 1.3645}, {"child": "MB-027", "qty": 0.0108, "unitCost": 130.0, "lineCost": 1.404}, {"child": "MB-039", "qty": 0.07344, "unitCost": 96.0381355932203, "lineCost": 7.053}, {"child": "MB-040", "qty": 0.0108, "unitCost": 114.00269541779, "lineCost": 1.2312}, {"child": "MB-042", "qty": 0.0108, "unitCost": 136.42723880597, "lineCost": 1.4734}, {"child": "MB-050", "qty": 0.0108, "unitCost": 114.0, "lineCost": 1.2312}, {"child": "MB-102", "qty": 0.0108, "unitCost": 265.0, "lineCost": 2.862}, {"child": "PPH-001", "qty": 2.37456, "unitCost": 67.4507739938081, "lineCost": 160.1659}]}, {"code": "SWT-002-24", "desc": "\u0642\u0627\u0644\u0628 \u062a\u062d\u0636\u064a\u0631 \u0645\u062a\u0639\u062f\u062f \u0627\u0644\u0627\u0633\u062a\u062e\u062f\u0627\u0645\u0627\u062a", "conv": 24, "sp": 1200.0, "rawBoxCost": 298.473, "boxCost": 402.9386, "profit": 797.0614, "profitPct": 66.42, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-SWT-002-24", "qty": 1.0, "unitCost": 16.0, "lineCost": 16.0}, {"child": "MB-017", "qty": 0.01056, "unitCost": 95.0, "lineCost": 1.0032}, {"child": "MB-019", "qty": 0.01056, "unitCost": 125.283018867925, "lineCost": 1.323}, {"child": "MB-021", "qty": 0.01056, "unitCost": 126.344086021505, "lineCost": 1.3342}, {"child": "MB-028", "qty": 0.01056, "unitCost": 130.0, "lineCost": 1.3728}, {"child": "MB-039", "qty": 0.0468, "unitCost": 96.0381355932203, "lineCost": 4.4946}, {"child": "MB-040", "qty": 0.01056, "unitCost": 114.00269541779, "lineCost": 1.2039}, {"child": "MB-042", "qty": 0.01056, "unitCost": 136.42723880597, "lineCost": 1.4407}, {"child": "PPH-001", "qty": 2.04864, "unitCost": 67.4507739938081, "lineCost": 138.1824}, {"child": "PPH-005", "qty": 1.5132, "unitCost": 55.5896038716616, "lineCost": 84.1182}, {"child": "ST-SWT-002-01", "qty": 24.0, "unitCost": 2.0, "lineCost": 48.0}]}, {"code": "PST-001-24", "desc": "\u0645\u0635\u0641\u0649 2 \u0641\u064a 1 (24)", "conv": 24, "sp": 792.0, "rawBoxCost": 0, "boxCost": 0.0, "profit": 0, "profitPct": 0, "hasBOM": false, "missing": [], "bomLines": []}, {"code": "ALM-101-03", "desc": "\u0637\u0642\u0645 \u0633\u0641\u0631\u0648 \u0623\u0644\u0645\u0648\u0646\u062f 13 \u0642\u0637\u0639\u0629", "conv": 3, "sp": 2250.0, "rawBoxCost": 708.5111, "boxCost": 956.49, "profit": 1293.51, "profitPct": 57.49, "hasBOM": true, "missing": [], "bomLines": [{"child": "CE-ALM-101-03", "qty": 1.0, "unitCost": 15.0, "lineCost": 15.0}, {"child": "CI-ALM-001-01", "qty": 3.0, "unitCost": 16.25, "lineCost": 48.75}, {"child": "SAN-001", "qty": 6.177, "unitCost": 104.380952380952, "lineCost": 644.7611}]}];

const _fmtNum = new Intl.NumberFormat("en-US-u-nu-latn", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const _fmtQtyNum = new Intl.NumberFormat("en-US-u-nu-latn", { minimumFractionDigits: 5, maximumFractionDigits: 5 });
const _fmtPct = new Intl.NumberFormat("en-US-u-nu-latn", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fmt = n => n == null ? undefined : _fmtNum.format(n);
const fmtQty = n => n == null ? undefined : _fmtQtyNum.format(n);
const fmtPct = n => _fmtPct.format(n);

const profitColor = p => p >= 60 ? "#15803d" : p >= 40 ? "#b45309" : p >= 20 ? "#c2410c" : "#dc2626";
const profitBg   = p => p >= 60 ? "#f0fdf4" : p >= 40 ? "#fffbeb" : p >= 20 ? "#fff7ed" : "#fef2f2";

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheets = {};
        wb.SheetNames.forEach(n => { sheets[n.trim()] = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: null }); });
        resolve(sheets);
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function processExcelData(sheets) {
  const bomSheet  = Object.entries(sheets).find(([k]) => k.toLowerCase().includes("bill") || k.toLowerCase() === "bom")?.[1];
  const costSheet = Object.entries(sheets).find(([k]) => k.toLowerCase().includes("cost") || k.toLowerCase().includes("raw"))?.[1];
  const priceSheet= Object.entries(sheets).find(([k]) => k.toLowerCase().includes("price"))?.[1];
  if (!bomSheet || !costSheet || !priceSheet) return null;

  const costMap = {};
  costSheet.forEach(r => { const c = String(r.ItemCode||"").trim(); const v = parseFloat(r.LastCost); if (c && !isNaN(v)) costMap[c] = v; });

  const bomMap = {};
  bomSheet.forEach(r => {
    const parent = String(r.ParentItemCode||"").trim();
    const child  = String(r.ChildItemCode||"").trim();
    const qty    = parseFloat(r.Quantity)||0;
    const bQty   = parseFloat(r.ParentBatchQty)||1;
    if (!parent) return;
    if (!bomMap[parent]) bomMap[parent] = { batchQty: bQty, lines: [] };
    bomMap[parent].lines.push({ child, qty });
  });

  const results = [];
  priceSheet.forEach(r => {
    const code = String(r.ItemCode||"").trim();
    const desc = String(r.ItemDescription||"").trim();
    const sp   = parseFloat(r.PriceSellingUnit)||0;
    const conv = parseFloat(r.SellingConversion)||1;
    const hasBOM = !!bomMap[code];

    let rawBoxCost = 0, missing = [], bomLines = [];
    if (hasBOM) {
      const { batchQty, lines } = bomMap[code];
      lines.forEach(({ child, qty }) => {
        const uc = costMap[child] ?? null;
        const lc = uc !== null ? qty * uc : 0;
        if (uc === null) missing.push(child);
        rawBoxCost += lc;
        bomLines.push({ child, qty, unitCost: uc, lineCost: lc, missing: uc === null });
      });
      rawBoxCost = (rawBoxCost / batchQty) * conv;
    }
    const boxCost = hasBOM ? Math.round(rawBoxCost * 1.35 * 100) / 100 : 0;
    rawBoxCost = Math.round(rawBoxCost * 100) / 100;
    const profit = hasBOM ? sp - boxCost : 0;
    const profitPct = hasBOM && sp > 0 ? Math.round((profit / sp * 100) * 10) / 10 : 0;
    results.push({ code, desc, conv: parseInt(conv), sp, rawBoxCost, boxCost, profit: Math.round(profit * 100) / 100, profitPct, hasBOM, missing, bomLines });
  });
  return results.sort((a, b) => a.code.localeCompare(b.code));
}

function BalanceGrid({ rows, showCostLabel, balSort, setBalSort, colFilters, setColFilters, fmt, fmtNum }) {
  const sorted = [...rows].sort((a, b) => {
    const { col, dir } = balSort;
    let va = a[col] ?? "", vb = b[col] ?? "";
    if (typeof va === "number" && typeof vb === "number") return dir === "asc" ? va - vb : vb - va;
    return dir === "asc" ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
  });
  const processed = sorted.filter(r =>
    Object.entries(colFilters).every(([c, v]) => !v || String(r[c]||"").toLowerCase().includes(v.toLowerCase()))
  );
  const tBal    = processed.reduce((s, r) => s + r.bal, 0);
  const tBoxes  = processed.reduce((s, r) => s + r.balInBoxes, 0);
  const tAmount = processed.filter(r => r.amount !== null).reduce((s, r) => s + r.amount, 0);
  const si = (col) => balSort.col === col ? (balSort.dir === "asc" ? "▲" : "▼") : "⇅";
  const onSort = (col) => setBalSort(s => ({ col, dir: s.col === col && s.dir === "asc" ? "desc" : "asc" }));

  const TH = ({ col, label, color, align = "right" }) => (
    <th style={{ padding: 0, borderBottom: "2px solid #e2e8f0", background: "#f8fafc", minWidth: 70 }}>
      <div onClick={() => onSort(col)}
        style={{ padding: "7px 10px 3px", display: "flex", alignItems: "center", justifyContent: align === "right" ? "flex-end" : align === "center" ? "center" : "flex-start", gap: 3, color, fontWeight: 700, fontSize: 11, cursor: "pointer", userSelect: "none", whiteSpace: "nowrap" }}>
        {label} <span style={{ fontSize: 9, opacity: balSort.col === col ? 1 : 0.3 }}>{si(col)}</span>
      </div>
      <div style={{ padding: "0 5px 5px" }}>
        <input
          value={colFilters[col] || ""}
          onChange={e => setColFilters(f => ({ ...f, [col]: e.target.value }))}
          placeholder="filter..."
          style={{ width: "100%", padding: "3px 5px", border: "1px solid #cbd5e1", borderRadius: 4, fontSize: 10, fontFamily: "inherit", background: "#fff", boxSizing: "border-box" }}
        />
      </div>
    </th>
  );

  return (
    <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 380px)" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
          <tr>
            <th style={{ padding: "9px 10px 14px", borderBottom: "2px solid #e2e8f0", background: "#f8fafc", color: "#94a3b8", textAlign: "left", fontSize: 11 }}>#</th>
            <TH col="ItemCode"         label="Item Code"      color="#475569" align="left" />
            <TH col="ItemDescription"  label="Description"    color="#475569" align="right" />
            <TH col="ItemType"         label="Type"           color="#64748b" align="center" />
            <TH col="Warehouse"        label="Warehouse"      color="#475569" align="left" />
            <TH col="bal"              label="Bal (Unit)"     color="#1d4ed8" />
            <TH col="balInBoxes"       label="Bal (Box)"      color="#0f2d5a" />
            <TH col="unitCost"         label={showCostLabel}  color="#7c3aed" />
            <TH col="amount"           label="Amount"         color="#0891b2" />
          </tr>
        </thead>
        <tbody>
          {processed.map((r, idx) => (
            <tr key={idx} style={{ borderBottom: "1px solid #f1f5f9", background: idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
              <td style={{ padding: "6px 10px", color: "#94a3b8", fontSize: 10 }}>{idx + 1}</td>
              <td style={{ padding: "6px 10px", fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>{r.ItemCode}</td>
              <td style={{ padding: "6px 10px", textAlign: "right", direction: "rtl", color: "#374151", fontSize: 11 }}>{r.ItemDescription || "—"}</td>
              <td style={{ padding: "6px 10px", textAlign: "center" }}>
                <span style={{ background: r.ItemType==="F"?"#eff6ff":r.ItemType==="R"?"#f0fdf4":"#fef9c3", color: r.ItemType==="F"?"#1d4ed8":r.ItemType==="R"?"#15803d":"#92400e", padding: "2px 6px", borderRadius: 5, fontSize: 10, fontWeight: 700 }}>{r.ItemType}</span>
              </td>
              <td style={{ padding: "6px 10px", fontSize: 11 }}>
                <span style={{ background: "#eff6ff", color: "#1d4ed8", padding: "2px 7px", borderRadius: 5, fontSize: 10, fontWeight: 600 }}>{r.Warehouse || "—"}</span>
              </td>
              <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 600, color: r.bal>0?"#15803d":r.bal<0?"#dc2626":"#94a3b8" }}>{fmtNum.format(r.bal)}</td>
              <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: r.balInBoxes>0?"#0f2d5a":r.balInBoxes<0?"#dc2626":"#94a3b8" }}>{fmtNum.format(r.balInBoxes)}</td>
              <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed" }}>{r.unitCost !== null ? fmt(r.unitCost) : <span style={{ color: "#94a3b8" }}>—</span>}</td>
              <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#0891b2" }}>{r.amount !== null ? fmt(r.amount) : <span style={{ color: "#94a3b8" }}>—</span>}</td>
            </tr>
          ))}
        </tbody>
        <tfoot style={{ position: "sticky", bottom: 0, background: "#f8fafc", borderTop: "2px solid #e2e8f0" }}>
          <tr>
            <td colSpan={5} style={{ padding: "9px 10px", fontWeight: 800, color: "#0f2d5a" }}>TOTAL ({processed.length} of {rows.length})</td>
            <td style={{ padding: "9px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#1d4ed8" }}>{fmtNum.format(Math.round(tBal))}</td>
            <td style={{ padding: "9px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#0f2d5a" }}>{fmtNum.format(Math.round(tBoxes * 1000) / 1000)}</td>
            <td style={{ padding: "9px 10px" }}></td>
            <td style={{ padding: "9px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#0891b2" }}>{fmt(tAmount)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

function Login({ onLogin }) {
  const [user, setUser] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");

  const handleSubmit = () => {
    if (user === "sysadmin" && pass === "sysadmin") {
      onLogin();
    } else {
      setError("Invalid username or password");
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f2d5a 0%, #1a4a8a 50%, #0f2d5a 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui, -apple-system, sans-serif" }}>
      <div style={{ background: "#fff", borderRadius: 18, boxShadow: "0 20px 60px rgba(0,0,0,0.3)", padding: "48px 40px", width: 360, textAlign: "center" }}>
        {/* Logo */}
        <div style={{ width: 72, height: 72, background: "linear-gradient(135deg, #0f2d5a, #1a6abf)", borderRadius: 18, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px", fontSize: 32 }}>⚙️</div>
        <div style={{ fontSize: 22, fontWeight: 800, color: "#0f2d5a", marginBottom: 4 }}>BOM</div>
        <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 32 }}>Pro Plast GLC — Sign in to continue</div>

        {/* Fields */}
        <div style={{ textAlign: "left", marginBottom: 16 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#475569", display: "block", marginBottom: 6 }}>USERNAME</label>
          <input value={user} onChange={e => { setUser(e.target.value); setError(""); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            placeholder="Enter username"
            style={{ width: "100%", padding: "11px 14px", border: "1.5px solid #e2e8f0", borderRadius: 9, fontSize: 14, fontFamily: "inherit", outline: "none", boxSizing: "border-box", transition: "border 0.2s" }}
            onFocus={e => e.target.style.borderColor = "#1a6abf"}
            onBlur={e => e.target.style.borderColor = "#e2e8f0"} />
        </div>
        <div style={{ textAlign: "left", marginBottom: 24 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#475569", display: "block", marginBottom: 6 }}>PASSWORD</label>
          <input value={pass} onChange={e => { setPass(e.target.value); setError(""); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            type="password" placeholder="Enter password"
            style={{ width: "100%", padding: "11px 14px", border: "1.5px solid #e2e8f0", borderRadius: 9, fontSize: 14, fontFamily: "inherit", outline: "none", boxSizing: "border-box", transition: "border 0.2s" }}
            onFocus={e => e.target.style.borderColor = "#1a6abf"}
            onBlur={e => e.target.style.borderColor = "#e2e8f0"} />
        </div>

        {error && (
          <div style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fecaca", borderRadius: 8, padding: "9px 14px", fontSize: 13, marginBottom: 16 }}>
            ❌ {error}
          </div>
        )}

        <button onClick={handleSubmit}
          style={{ width: "100%", padding: "13px", background: "linear-gradient(135deg, #0f2d5a, #1a6abf)", color: "#fff", border: "none", borderRadius: 10, fontSize: 15, fontWeight: 700, fontFamily: "inherit", cursor: "pointer", letterSpacing: 0.5 }}>
          Sign In →
        </button>

        <div style={{ marginTop: 24, fontSize: 11, color: "#cbd5e1" }}>Pro Plast GLC © 2026</div>
      </div>
    </div>
  );
}

const INITIAL_SALES_DATA = [];

function SalesSTH({ col, children, align="right", width=70, saSort, setSaSort, saColF, setSaColF }) {
  return (
    <th style={{ padding:0, borderBottom:"2px solid #e2e8f0", background:"#f8fafc", minWidth:width }}>
      <div onClick={()=>setSaSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}
        style={{ padding:"6px 8px 2px", display:"flex", alignItems:"center", justifyContent:align==="right"?"flex-end":align==="center"?"center":"flex-start", gap:3, color:saSort.col===col?"#1d4ed8":"#0f2d5a", fontWeight:700, fontSize:11, cursor:"pointer", userSelect:"none", whiteSpace:"nowrap" }}>
        {children}<span style={{ fontSize:9, opacity:saSort.col===col?1:0.35 }}>{saSort.col===col?(saSort.dir==="desc"?"▼":"▲"):"⇅"}</span>
      </div>
      <div style={{ padding:"0 4px 4px" }}>
        <input value={saColF[col]||""} onChange={e=>setSaColF(f=>({...f,[col]:e.target.value}))}
          placeholder="filter..." style={{ width:"100%", padding:"2px 4px", border:"1px solid #cbd5e1", borderRadius:3, fontSize:10, fontFamily:"inherit", background:saColF[col]?"#eff6ff":"#fff", boxSizing:"border-box" }} />
      </div>
    </th>
  );
}

const API_URL   = "https://sila.silasystem.com:7103/General/GeneralAPI/";
const HEADERS   = { "Accept": "application/json", "content-type": "application/json", "Sp_Name": "APIClaudeOperation" };
const BASE_BODY = { User: "mhd", AppVersionWeb: "225", AppVersionAndroid: "225", AppVersionIos: "225", AppVersionDesktop: "225", FireBaseToken: "", PlatForm: "web", deviceID: "", IP: "192.168.1.3" };

/* ═══ Cash Flow helper components (module-level to avoid React reconciler issues) ═══ */
function CFSectionTitle({ icon, title, accent = "#0f2d5a" }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "28px 0 14px", paddingBottom: 8, borderBottom: `2px solid ${accent}22` }}>
      <div style={{ width: 32, height: 32, borderRadius: 8, background: `${accent}12`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>{icon}</div>
      <span style={{ fontSize: 15, fontWeight: 800, color: accent, letterSpacing: 0.3 }}>{title}</span>
    </div>
  );
}
function CFMetricCard({ label, value, sub, color = "#0f2d5a", trend, icon }) {
  const n = (v) => parseFloat(v) || 0;
  return (
    <div style={{ background: "#fff", borderRadius: 12, padding: "18px 20px", border: "1px solid #e2e8f0", position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 3, background: color, opacity: 0.7 }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 1, textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
          <div style={{ fontSize: 22, fontWeight: 900, fontFamily: "monospace", color, lineHeight: 1.1 }}>{value}</div>
          {sub && <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 6 }}>{sub}</div>}
        </div>
        {icon && <div style={{ width: 36, height: 36, borderRadius: 10, background: `${color}10`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>{icon}</div>}
      </div>
      {trend !== undefined && (
        <div style={{ display: "inline-flex", alignItems: "center", gap: 4, marginTop: 8, padding: "2px 8px", borderRadius: 20, background: n(trend) >= 0 ? "#dcfce7" : "#fef2f2", fontSize: 10, fontWeight: 700, color: n(trend) >= 0 ? "#15803d" : "#dc2626" }}>
          {n(trend) >= 0 ? "▲" : "▼"} {Math.abs(n(trend)).toFixed(2)}%
        </div>
      )}
    </div>
  );
}
function CFGrid({ cols = 3, children }) {
  return <div style={{ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 12 }}>{children}</div>;
}
function CFMiniBar({ items }) {
  const max = Math.max(...items.map(i => Math.abs(i.value)), 1);
  const fmt2 = (v) => new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(parseFloat(v) || 0);
  return (
    <div style={{ background: "#fff", borderRadius: 12, padding: "18px 20px", border: "1px solid #e2e8f0" }}>
      {items.map((it, i) => (
        <div key={i} style={{ marginBottom: i < items.length - 1 ? 12 : 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
            <span style={{ fontSize: 11, fontWeight: 600, color: "#64748b" }}>{it.label}</span>
            <span style={{ fontSize: 12, fontWeight: 800, fontFamily: "monospace", color: it.color || "#0f2d5a" }}>{fmt2(it.value)}</span>
          </div>
          <div style={{ height: 6, borderRadius: 3, background: "#f1f5f9", overflow: "hidden" }}>
            <div style={{ height: "100%", borderRadius: 3, background: it.color || "#1d4ed8", width: `${(Math.abs(it.value) / max) * 100}%`, transition: "width 0.5s ease" }} />
          </div>
        </div>
      ))}
    </div>
  );
}
function CFVSGauge({ leftVal, rightVal, leftLabel, rightLabel, leftColor = "#16a34a", rightColor = "#dc2626" }) {
  const fmt0 = (v) => new Intl.NumberFormat("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(parseFloat(v) || 0);
  const total = Math.abs(leftVal) + Math.abs(rightVal) || 1;
  const leftPct = (Math.abs(leftVal) / total * 100).toFixed(0);
  return (
    <div style={{ background: "#fff", borderRadius: 14, padding: "24px", border: "1px solid #e2e8f0" }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 16 }}>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: leftColor, letterSpacing: 1, marginBottom: 4 }}>{leftLabel}</div>
          <div style={{ fontSize: 28, fontWeight: 900, fontFamily: "monospace", color: leftColor }}>{fmt0(leftVal)}</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ width: 44, height: 44, borderRadius: 22, background: "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 13, color: "#64748b" }}>VS</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: rightColor, letterSpacing: 1, marginBottom: 4 }}>{rightLabel}</div>
          <div style={{ fontSize: 28, fontWeight: 900, fontFamily: "monospace", color: rightColor }}>{fmt0(rightVal)}</div>
        </div>
      </div>
      <div style={{ height: 10, borderRadius: 5, background: rightColor + "22", overflow: "hidden" }}>
        <div style={{ height: "100%", borderRadius: 5, background: `linear-gradient(90deg, ${leftColor}, ${leftColor}cc)`, width: `${leftPct}%`, transition: "width 0.5s ease" }} />
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <span style={{ fontSize: 10, fontWeight: 700, color: leftColor }}>{leftPct}%</span>
        <span style={{ fontSize: 10, fontWeight: 700, color: rightColor }}>{100 - parseInt(leftPct)}%</span>
      </div>
    </div>
  );
}

export default function App() {
  const [loggedIn, setLoggedIn] = useState(() => { try { return sessionStorage.getItem("cm_auth") === "1"; } catch(e) { return false; } });
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // ── Auto-logout after 60 min inactivity ──
  const TIMEOUT_MS = 60 * 60 * 1000; // 60 minutes
  const [autoLogout, setAutoLogout] = useState(true);
  const timerRef = useRef(null);
  const [countdown, setCountdown] = useState(null); // shows warning in last 60s

  const doLogout = useCallback(() => {
    try { sessionStorage.removeItem("cm_auth"); } catch(e) { void 0; }
    setLoggedIn(false);
    setCountdown(null);
  }, []);

  const resetTimer = useCallback(() => {
    if (!autoLogout) return;
    clearTimeout(timerRef.current);
    setCountdown(null);
    // Warning at 59 min
    timerRef.current = setTimeout(() => {
      let secs = 60;
      setCountdown(secs);
      const tick = setInterval(() => {
        secs -= 1;
        if (secs <= 0) { clearInterval(tick); doLogout(); }
        else setCountdown(secs);
      }, 1000);
      timerRef.current = tick;
    }, TIMEOUT_MS - 60000);
  }, [autoLogout, doLogout]);

  useEffect(() => {
    if (!loggedIn || !autoLogout) { clearTimeout(timerRef.current); setCountdown(null); return; }
    const events = ["mousemove", "mousedown", "keydown", "touchstart", "scroll", "click"];
    events.forEach(e => window.addEventListener(e, resetTimer, { passive: true }));
    resetTimer();
    return () => {
      events.forEach(e => window.removeEventListener(e, resetTimer));
      clearTimeout(timerRef.current);
    };
  }, [loggedIn, autoLogout, resetTimer]);

  const [page, setPage]         = useState("cost");
  const [data, setData]         = useState(INITIAL_DATA);
  const [costOverrides, setCostOverrides] = useState({}); // { itemCode: newCost }
  const [editingCode, setEditingCode] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [search, setSearch]     = useState("");
  const [filterBOM, setFilter]  = useState("all");
  const [rmSearch, setRmSearch] = useState("");
  const [rmSort, setRmSort]     = useState("code");
  const [rmDir, setRmDir]       = useState("asc");
  const [selectedRaw, setSelectedRaw] = useState(null);
  const [fgSearch, setFgSearch] = useState("");
  const [fgFilter, setFgFilter] = useState("all");
  const [balanceData, setBalanceData]   = useState([]);
  const [balanceLoading, setBalanceLoading] = useState(false);
  const [balanceMsg, setBalanceMsg]     = useState("");
  const [balanceSearch, setBalanceSearch] = useState("");
  const [balanceWarehouse, setBalanceWarehouse] = useState("all");
  const [balanceItemType, setBalanceItemType] = useState("all");
  const [balanceTab, setBalanceTab]     = useState("fg");
  const [apiRawCost, setApiRawCost]     = useState([]);
  const [balSort, setBalSort]           = useState({ col: "ItemCode", dir: "asc" });
  const [colFilters, setColFilters]     = useState({});
  const [cfData, setCfData]             = useState(null);
  const [cfLoading, setCfLoading]       = useState(false);
  const [cfMsg, setCfMsg]               = useState("");

  const [plData, setPlData]             = useState(null);
  const [plLoading, setPlLoading]       = useState(false);
  const [plMsg, setPlMsg]               = useState("");
  const [plSearch, setPlSearch]         = useState("");
  const [plPriceType, setPlPriceType]   = useState("");
  const [plDiscount, setPlDiscount]     = useState(25);

  // ── Sales Analysis ──
  const [saData, setSaData]             = useState(INITIAL_SALES_DATA);
  const [saLoading, setSaLoading]       = useState(false);
  const [saMsg, setSaMsg]               = useState("");
  const [saYear, setSaYear]             = useState(new Date().getFullYear());
  const [saView, setSaView]             = useState(["item"]);
  const [saSearch, setSaSearch]         = useState("");
  const [saSort, setSaSort]             = useState({ col: "amount", dir: "desc" });
  const [saOpenMonth, setSaOpenMonth]   = useState(false);
  const [saOpenGroup, setSaOpenGroup]   = useState(false);
  const [saOpenCustomer, setSaOpenCustomer] = useState(false);
  const [saColF, setSaColF]             = useState({});
  const [saFilter, setSaFilter]         = useState({});  // { item: ["code1",...], customer: ["no1",...], salesperson: [...], month: ["1","3",...] }
  const [saItemSearch, setSaItemSearch] = useState("");
  const [saCustSearch, setSaCustSearch] = useState("");
  const [saFamilySearch, setSaFamilySearch] = useState("");
  const [saSelected, setSaSelected]     = useState(null); // { type: "item"|"customer", key, label }

  // ── Price Lab ──
  const [plabPriceType, setPlabPriceType] = useState("");
  const [plabSearch, setPlabSearch]       = useState("");
  const [plabPrices, setPlabPrices]       = useState({}); // { ItemCode: newPrice }
  const [plabDiscount, setPlabDiscount]   = useState(0);
  const [plabSelected, setPlabSelected]   = useState(null); // ItemCode

  // ── Customer Coverage ──
  const [ccData, setCcData]         = useState(null);
  const [ccLoading, setCcLoading]   = useState(false);
  const [ccMsg, setCcMsg]           = useState("");
  const [ccYear, setCcYear]         = useState(new Date().getFullYear());
  const [ccTab, setCcTab]           = useState("customer"); // "customer" | "salesperson"
  const [ccSearch, setCcSearch]     = useState("");
  const [ccMonth, setCcMonth]       = useState(0); // 0 = all months
  const [ccSort, setCcSort]         = useState({ col:"coverage", dir:"desc" });

  // ── Expenses Analysis ──
  const [exData,         setExData]         = useState([]);
  const [exLoading,      setExLoading]      = useState(false);
  const [exMsg,          setExMsg]          = useState("");
  const [exYear,         setExYear]         = useState(new Date().getFullYear());
  const [exGroupBy,      setExGroupBy]      = useState(["account"]);
  const [exSearch,       setExSearch]       = useState("");
  const [exMonth,        setExMonth]        = useState([]);
  const [exParent,       setExParent]       = useState([]);
  const [exParentSearch, setExParentSearch] = useState("");
  const [exSort,         setExSort]         = useState({ col: "Amount", dir: "desc" });
  const [exSelected,     setExSelected]     = useState(null);
  const [exDetailTab,    setExDetailTab]    = useState("monthly");
  const [exParentOpen,   setExParentOpen]   = useState(false);
  const [exMonthOpen,    setExMonthOpen]    = useState(false);
  const [exGroupOpen,    setExGroupOpen]    = useState(false);

  // Recalculate all BOM items whenever costOverrides changes
  const recalcData = useMemo(() => {
    if (Object.keys(costOverrides).length === 0) return data;
    return data.map(item => {
      if (!item.hasBOM) return item;
      const newLines = item.bomLines.map(l => {
        const uc = costOverrides[l.child] !== undefined ? costOverrides[l.child] : l.unitCost;
        return { ...l, unitCost: uc, lineCost: uc !== null ? Math.round(l.qty * uc * 10000) / 10000 : 0 };
      });
      const newRaw = Math.round(newLines.reduce((s, l) => s + l.lineCost, 0) * 100) / 100;
      const newBox = Math.round(newRaw * 1.35 * 100) / 100;
      const newProfit = Math.round((item.sp - newBox) * 100) / 100;
      const newPct = item.sp > 0 ? Math.round((newProfit / item.sp * 100) * 10) / 10 : 0;
      return { ...item, bomLines: newLines, rawBoxCost: newRaw, boxCost: newBox, profit: newProfit, profitPct: newPct };
    });
  }, [data, costOverrides]);

  // Compute which parent items are affected by current overrides
  const affectedItems = useMemo(() => {
    if (Object.keys(costOverrides).length === 0) return [];
    return data
      .filter(item => item.hasBOM && item.bomLines.some(l => costOverrides[l.child] !== undefined))
      .map(item => {
        const newItem = recalcData.find(r => r.code === item.code);
        return {
          code: item.code,
          desc: item.desc,
          oldRaw: item.rawBoxCost,
          newRaw: newItem.rawBoxCost,
          oldBox: item.boxCost,
          newBox: newItem.boxCost,
          oldProfit: item.profit,
          newProfit: newItem.profit,
          oldPct: item.profitPct,
          newPct: newItem.profitPct,
          sp: item.sp,
          changedChildren: item.bomLines.filter(l => costOverrides[l.child] !== undefined).map(l => l.child),
        };
      })
      .sort((a, b) => (b.newPct - b.oldPct) - (a.newPct - a.oldPct)); // biggest losers first
  }, [data, recalcData, costOverrides]);

  const handleSaveOverride = (code) => {
    const val = parseFloat(editValue);
    if (!isNaN(val) && val >= 0) {
      setCostOverrides(prev => ({ ...prev, [code]: val }));
    }
    setEditingCode(null);
    setEditValue("");
  };

  const handleResetOverride = (code) => {
    setCostOverrides(prev => { const n = { ...prev }; delete n[code]; return n; });
  };
  const [sortBy, setSortBy]     = useState("code");
  const [sortDir, setSortDir]   = useState("asc");
  const [selected, setSelected] = useState(null);
  const [tab, setTab]           = useState("bom");
  const [discount, setDiscount]   = useState(25);
  const [scenario, setScenario] = useState({});
  const [uploadMsg, setMsg]       = useState("");
  const [uploading, setUploading] = useState(false);
  const [apiUrl, setApiUrl]       = useState(() => { try { return localStorage.getItem("cm_api_url") || ""; } catch(e) { return ""; } });
  const [refreshing, setRefreshing] = useState(false);
  const [refreshMsg, setRefreshMsg] = useState("");
  const [testData, setTestData]     = useState(null);
  const [testing, setTesting]       = useState(false);
  const fileRef = useRef();

  const handleUpload = async e => {
    const file = e.target.files[0]; if (!file) return;
    setUploading(true); setMsg("");
    try {
      const sheets = await parseExcel(file);
      const result = processExcelData(sheets);
      if (!result) { setMsg("❌ Missing sheets (Bill Of Material, Raw Cost, Price List)"); }
      else { setData(result); setMsg(`✅ Loaded ${result.length} items`); setSelected(null); }
    } catch (err) { setMsg("❌ " + err.message); }
    setUploading(false); e.target.value = "";
  };

  const handleRefresh = async () => {
    setRefreshing(true); setRefreshMsg("");
    try {
      const res = await fetch("https://sila.silasystem.com:7103/General/GeneralAPI/", {
        method: "POST",
        headers: {
          "Accept":         "application/json",
          "content-type":   "application/json",
          "Sp_Name":        "APIClaudeOperation"
        },
        body: JSON.stringify({
          Operation:         "Get All Data",
          User:              "mhd",
          AppVersionWeb:     "225",
          AppVersionAndroid: "225",
          AppVersionIos:     "225",
          AppVersionDesktop: "225",
          FireBaseToken:     "",
          PlatForm:          "web",
          deviceID:          "",
          IP:                "192.168.1.3"
        })
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();

      // Response structure:
      // datasets[0] = Finished Goods    (ItemCode, ItemDescription, SellingConversion)
      // datasets[1] = Price List        (ItemID, ItemCode, ItemDescription, PriceSellingUnit, SellingUM, SellingConversion)
      // datasets[2] = Raw Cost          (ItemID, ItemCode, ItemDescription, LastCost, AverageCost, ItemType)
      // datasets[3] = BOM               (ParentItemCode, Line, ChildItemCode, Quantity, ParentBatchQty)
      // datasets[4] = Sales Amount      (ItemCode, Qty, Amount)
      // Response uses List0, List1, List2, List3, List4
      const fgList      = json.List0 || [];
      const priceList   = json.List1 || [];
      const rawCostList = json.List2 || [];
      const bomList     = json.List3 || [];
      const salesList   = json.List4 || [];

      if (!fgList.length && !priceList.length) throw new Error("API returned empty data — check SP_Name header");

      // FG descriptions map
      const fgDescMap = Object.fromEntries(
        fgList.map(r => [String(r.ItemCode||"").trim(), String(r.ItemDescription||"").trim()])
      );

      // Price map — ItemCode → { PriceSellingUnit, SellingConversion, ItemDescription }
      const priceMap = Object.fromEntries(
        priceList.map(r => [String(r.ItemCode||"").trim(), r])
      );

      // Merged price sheet: all FG items with price data where available
      const allCodes = [...new Set([
        ...fgList.map(r => String(r.ItemCode||"").trim()),
        ...priceList.map(r => String(r.ItemCode||"").trim())
      ])];

      const mergedPriceSheet = allCodes.map(code => {
        const p    = priceMap[code] || {};
        const desc = fgDescMap[code] || String(p.ItemDescription||"").trim();
        // SellingConversion: from price list first, else from FG list
        const fgConv = fgList.find(r => String(r.ItemCode||"").trim() === code)?.SellingConversion;
        const conv   = p.SellingConversion ?? fgConv ?? null;
        return {
          ItemCode:          code,
          ItemDescription:   desc,
          PriceSellingUnit:  p.PriceSellingUnit  ?? 0,
          SellingConversion: conv,
        };
      });

      const sheets = {
        "Price List": mergedPriceSheet,
        "Raw Cost":   rawCostList.map(r => ({
          ItemCode:    String(r.ItemCode||"").trim(),
          AverageCost: r.AverageCost,
          LastCost:    r.LastCost,
          ItemType:    r.ItemType
        })),
        "Bill Of Material": bomList.map(r => ({
          ParentItemCode: String(r.ParentItemCode||"").trim(),
          ChildItemCode:  String(r.ChildItemCode||"").trim(),
          Quantity:       r.Quantity,
          ParentBatchQty: r.ParentBatchQty
        }))
      };

      const result = processExcelData(sheets);
      if (!result) throw new Error("Could not process API data — check SP field names");

      setData(result);
      setCostOverrides({});
      setSelected(null);
      setApiRawCost(rawCostList);
      // Save full price list (all types) for Price List page
      if (priceList.length > 0) setPlData({ List0: priceList });

      // Update Sales Amount data if returned
      if (salesList.length > 0) {
        window.__apiSalesData = salesList.map(r => ({
          code:   String(r.ItemCode||"").trim(),
          qty:    parseFloat(r.Qty||0),
          amount: parseFloat(r.Amount||0)
        }));
      }

      setRefreshMsg(`✅ Refreshed — ${result.length} items · ${fgList.length} FG · ${rawCostList.length} materials · ${bomList.length} BOM lines · ${salesList.length} sales rows`);
    } catch (err) {
      setRefreshMsg("❌ " + err.message);
    }
    setRefreshing(false);
  };

  const handleTest = async () => {
    setTesting(true); setTestData(null);
    try {
      const res = await fetch("https://sila.silasystem.com:7103/General/GeneralAPI/", {
        method: "POST",
        headers: {
          "Accept":         "application/json",
          "content-type":   "application/json",
          "Sp_Name":        "APIClaudeOperation"
        },
        body: JSON.stringify({
          Operation:         "Get All Data",
          User:              "mhd",
          AppVersionWeb:     "225",
          AppVersionAndroid: "225",
          AppVersionIos:     "225",
          AppVersionDesktop: "225",
          FireBaseToken:     "",
          PlatForm:          "web",
          deviceID:          "",
          IP:                "192.168.1.3"
        })
      });
      const json = await res.json();
      setTestData({ ok: true, status: res.status, data: json });
    } catch (err) {
      setTestData({ ok: false, error: err.message });
    }
    setTesting(false);
  };

  const fetchBalance = async () => {
    setBalanceLoading(true); setBalanceMsg("");
    try {
      // Always fetch balance
      const res = await fetch(API_URL, { method: "POST", headers: HEADERS, body: JSON.stringify({ ...BASE_BODY, Operation: "Item Balance" }) });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const rows = json.List0 || json.List || (Array.isArray(json) ? json : []);
      setBalanceData(rows);

      // If apiRawCost not loaded yet, also fetch Get All Data to get List2
      if (apiRawCost.length === 0) {
        const res2 = await fetch(API_URL, { method: "POST", headers: HEADERS, body: JSON.stringify({ ...BASE_BODY, Operation: "Get All Data" }) });
        if (res2.ok) {
          const json2 = await res2.json();
          if (json2.List2?.length > 0) setApiRawCost(json2.List2);
        }
      }

      setBalanceMsg(`✅ Loaded ${rows.length} items`);
    } catch (err) {
      setBalanceMsg("❌ " + err.message);
    }
    setBalanceLoading(false);
  };

  const fetchCashFlow = async () => {
    setCfLoading(true); setCfMsg(""); setCfData(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST", headers: HEADERS,
        body: JSON.stringify({ ...BASE_BODY, Operation: "Cash Flow" })
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      setCfData(json);
      setCfMsg("✅ Loaded");
    } catch (err) {
      setCfMsg("❌ " + err.message);
    }
    setCfLoading(false);
  };

  const fetchPriceList = async () => {
    setPlLoading(true); setPlMsg(""); setPlData(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST", headers: HEADERS,
        body: JSON.stringify({ ...BASE_BODY, Operation: "Price List Type" })
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : (json?.List0 || json?.List || []);
      setPlData({ List0: rows });
      setPlMsg("✅ Loaded " + rows.length + " rows");
    } catch (err) {
      setPlMsg("❌ " + err.message);
    }
    setPlLoading(false);
  };

  const fetchCoverage = async (year) => {    setCcLoading(true); setCcMsg(""); setCcData(null);
    try {
      const res = await fetch(API_URL, { method:"POST", headers:HEADERS, body:JSON.stringify({ ...BASE_BODY, Operation:"Customer Coverage", Year:String(year||new Date().getFullYear()) }) });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const customers    = json?.List0 || [];
      const salespersons = json?.List1 || [];
      setCcData({ customers, salespersons });
      setCcMsg(`✅ ${customers.length} customers · ${salespersons.length} salespersons`);
    } catch(err) {
      setCcMsg("❌ " + err.message);
    }
    setCcLoading(false);
  };

  const fetchExpenses = async (year) => {
    setExLoading(true); setExMsg(""); setExData([]); setExSelected(null);
    try {
      const res = await fetch(API_URL, {
        method: "POST", headers: HEADERS,
        body: JSON.stringify({ ...BASE_BODY, Operation: "Expenses", Year: String(year || new Date().getFullYear()) })
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : (json?.List0 || json?.List || []);
      window.__exDebugRows = rows;
      if (rows.length > 0) {
        const firstRow = rows[0];
        const amtVal = firstRow.Amount ?? firstRow.amount ?? firstRow.NetAmount ?? "NOT FOUND";
        const amtType = typeof amtVal;
        const parsed = parseFloat(amtVal);
        setExMsg(`✅ ${rows.length} rows · Amount[0]="${amtVal}" (${amtType}) → parsed=${parsed} · Keys: ${Object.keys(firstRow).join(", ")}`);
      } else {
        setExMsg("No data found.");
      }
      setExData(rows);
    } catch (err) {
      setExMsg("❌ " + err.message);
    }
    setExLoading(false);
  };

  const fetchSales = async (year) => {
    try {
      const res = await fetch(API_URL, { method: "POST", headers: HEADERS, body: JSON.stringify({ ...BASE_BODY, Operation: "Sales Analysis", Year: String(year) }) });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : (json?.List0 || json?.List || []);
      setSaData(rows.length > 0 ? rows : INITIAL_SALES_DATA);
      setSaMsg(`✅ Loaded ${rows.length > 0 ? rows.length : INITIAL_SALES_DATA.length} rows${rows.length === 0 ? " (sample)" : ""}`);
    } catch (err) {
      setSaMsg("❌ " + err.message);
    }
    setSaLoading(false);
  };

  const filtered = useMemo(() => {
    let d = [...recalcData];
    if (search) d = d.filter(i => i.code.toLowerCase().includes(search.toLowerCase()) || i.desc.includes(search));
    if (filterBOM === "with")    d = d.filter(i => i.hasBOM);
    if (filterBOM === "without") d = d.filter(i => !i.hasBOM);
    if (filterBOM === "partial") d = d.filter(i => i.hasBOM && i.missing.length > 0);
    return d.sort((a, b) => {
      let av = a[sortBy], bv = b[sortBy];
      if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      return sortDir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
    });
  }, [data, search, filterBOM, sortBy, sortDir]);

  const stats = useMemo(() => {
    const wb = recalcData.filter(i => i.hasBOM);
    return { total: data.length, withBOM: wb.length, noBOM: data.filter(i => !i.hasBOM).length, partial: data.filter(i => i.hasBOM && i.missing.length > 0).length, avgProfit: wb.length ? wb.reduce((s, i) => s + i.profitPct, 0) / wb.length : 0, low: wb.filter(i => i.profitPct < 40).length };
  }, [data]);

  const selItem = selected ? recalcData.find(i => i.code === selected) : null;

  const scenarioResult = useMemo(() => {
    if (!selItem?.hasBOM) return null;
    const newBatchTotal = selItem.bomLines.reduce((s, l) => {
      const uc = scenario[l.child] !== undefined ? (parseFloat(scenario[l.child]) || 0) : (l.unitCost || 0);
      return s + l.qty * uc;
    }, 0);
    const ratio = selItem.rawBoxCost > 0 ? selItem.boxCost / selItem.rawBoxCost : 1.35;
    const newBoxCost = Math.round(newBatchTotal * ratio * 100) / 100;
    const newProfit  = Math.round((selItem.sp - newBoxCost) * 100) / 100;
    const newPct     = selItem.sp > 0 ? Math.round((newProfit / selItem.sp * 100) * 10) / 10 : 0;
    return { newBoxCost, newProfit, newPct };
  }, [selItem, scenario]);

  const affectedSet = useMemo(() => new Set(affectedItems.map(i => i.code)), [affectedItems]);
  const affectedMap = useMemo(() => Object.fromEntries(affectedItems.map(i => [i.code, i])), [affectedItems]);

  const sortBtn = (col, label) => (
    <button onClick={() => { setSortBy(col); setSortDir(d => d === "asc" && sortBy === col ? "desc" : "asc"); }}
      style={{ background: "none", border: "none", cursor: "pointer", fontWeight: sortBy === col ? 700 : 500, color: sortBy === col ? "#1d4ed8" : "#374151", fontSize: 12, fontFamily: "inherit", whiteSpace: "nowrap" }}>
      {label}{sortBy === col ? (sortDir === "asc" ? " ↑" : " ↓") : ""}
    </button>
  );

  return (
    <>
    {!loggedIn && <Login onLogin={() => { try { sessionStorage.setItem("cm_auth", "1"); } catch(e) { void 0; } setLoggedIn(true); }} />}
    {loggedIn && <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", display: "flex", minHeight: "100vh", direction: "ltr" }}>

      {/* ── Responsive CSS ── */}
      <style>{`
        * { box-sizing: border-box; }
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        input[type=number] { -moz-appearance: textfield; }
        :root { --sidebar-offset: 240px; }
        .cm-sidebar {
          transform: translateX(0);
          transition: transform 0.25s ease;
        }
        .cm-overlay {
          display: none;
          position: fixed;
          inset: 0;
          background: rgba(0,0,0,0.5);
          z-index: 99;
        }
        .cm-hamburger {
          display: none;
          align-items: center;
          justify-content: center;
          width: 36px;
          height: 36px;
          background: none;
          border: 1px solid #e2e8f0;
          border-radius: 8px;
          cursor: pointer;
          flex-shrink: 0;
        }
        @media (max-width: 768px) {
          :root { --sidebar-offset: 0px; }
          .cm-sidebar {
            transform: translateX(-100%) !important;
          }
          .cm-sidebar.open {
            transform: translateX(0) !important;
          }
          .cm-sidebar.open ~ .cm-overlay,
          .cm-overlay.active {
            display: block !important;
          }
          .cm-main {
            margin-left: 0 !important;
          }
          .cm-hamburger {
            display: flex !important;
          }
          .cm-topbar-title {
            font-size: 15px !important;
          }
        }
      `}</style>

      {/* ── Overlay (mobile) ── */}
      <div className={`cm-overlay${sidebarOpen ? " active" : ""}`} onClick={() => setSidebarOpen(false)} />

      {/* ── Sidebar ── */}
      <div className={`cm-sidebar${sidebarOpen ? " open" : ""}`} style={{ width: 240, minWidth: 240, background: "linear-gradient(180deg, #0f2d5a 0%, #1a3a6e 100%)", display: "flex", flexDirection: "column", position: "fixed", top: 0, left: 0, height: "100vh", zIndex: 100, boxShadow: "4px 0 20px rgba(0,0,0,0.15)" }}>

        {/* Logo */}
        <div style={{ padding: "24px 20px 20px", borderBottom: "1px solid rgba(255,255,255,0.1)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 40, height: 40, background: "linear-gradient(135deg, #1d4ed8, #3b82f6)", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>⚙️</div>
            <div style={{ flex: 1 }}>
              <div style={{ color: "#fff", fontWeight: 800, fontSize: 14, lineHeight: 1.2 }}>BOM</div>
              <div style={{ color: "rgba(255,255,255,0.5)", fontSize: 11, marginTop: 2 }}>Pro Plast GLC</div>
            </div>
            <button className="cm-hamburger" onClick={() => setSidebarOpen(false)} style={{ color: "rgba(255,255,255,0.7)", border: "1px solid rgba(255,255,255,0.2)", borderRadius: 6, width: 28, height: 28, fontSize: 16, cursor: "pointer", background: "none" }}>✕</button>
          </div>
        </div>

        {/* Nav Items */}
        <nav style={{ flex: 1, padding: "16px 12px", overflowY: "auto" }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: "rgba(255,255,255,0.35)", letterSpacing: 1.5, marginBottom: 8, paddingLeft: 8 }}>NAVIGATION</div>
          {[
            ["cost",      "📊", "BOM"],
            ["raw",       "🧪", "Raw Materials"],
            ["salesamt", "💰", "Sales Amount"],
            ["fg",        "🏷️", "Finished Goods"],
            ["balance",   "📦", "Item Balance"],
            ["rawcost",   "💎", "Raw Material Cost"],
            ["cashflow",  "💵", "Cash Flow"],
            ["pricelist", "🏷️", "Price List Type"],
            ["salesan",   "📈", "Sales Analysis"],
            ["pricelab",  "🔬", "Price Lab"],
            ["coverage",  "📊", "Customer Coverage"],
            ["expenses",  "💸", "Expenses Analysis"],
          ].map(([p, icon, label]) => (
            <button key={p} onClick={() => { setPage(p); setSidebarOpen(false); }} style={{
              width: "100%", display: "flex", alignItems: "center", gap: 12,
              padding: "11px 14px", borderRadius: 10, border: "none", cursor: "pointer",
              fontFamily: "inherit", fontSize: 13, fontWeight: 600, marginBottom: 4,
              background: page === p ? "rgba(255,255,255,0.15)" : "transparent",
              color: page === p ? "#fff" : "rgba(255,255,255,0.6)",
              borderLeft: page === p ? "3px solid #60a5fa" : "3px solid transparent",
              transition: "all 0.15s",
            }}>
              <span style={{ fontSize: 16 }}>{icon}</span>
              {label}
            </button>
          ))}
        </nav>

        {/* Bottom Actions */}
        <div style={{ padding: "12px", borderTop: "1px solid rgba(255,255,255,0.1)" }}>
          <button onClick={handleRefresh} disabled={refreshing} style={{
            width: "100%", padding: "9px", background: refreshing ? "rgba(255,255,255,0.05)" : "#15803d",
            border: "none", color: "#fff", borderRadius: 8, cursor: refreshing ? "not-allowed" : "pointer",
            fontSize: 12, fontWeight: 700, fontFamily: "inherit", marginBottom: 6
          }}>{refreshing ? "⏳ Loading..." : "🔄 Refresh DB"}</button>

          <button onClick={handleTest} disabled={testing} style={{
            width: "100%", padding: "9px", background: testing ? "rgba(255,255,255,0.05)" : "#7c3aed",
            border: "none", color: "#fff", borderRadius: 8, cursor: testing ? "not-allowed" : "pointer",
            fontSize: 12, fontWeight: 700, fontFamily: "inherit", marginBottom: 6
          }}>{testing ? "⏳ Testing..." : "🧪 Test API"}</button>

          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={() => fileRef.current.click()} disabled={uploading} style={{
              flex: 1, padding: "8px", background: "rgba(255,255,255,0.1)",
              border: "1px solid rgba(255,255,255,0.2)", color: "#fff", borderRadius: 8,
              cursor: "pointer", fontSize: 11, fontWeight: 600, fontFamily: "inherit"
            }}>{uploading ? "⏳" : "📂 Upload"}</button>
            <button onClick={() => { try { sessionStorage.removeItem("cm_auth"); } catch(e) { void 0; } setLoggedIn(false); }} style={{
              padding: "8px 12px", background: "rgba(255,255,255,0.1)",
              border: "1px solid rgba(255,255,255,0.2)", color: "#fff", borderRadius: 8,
              cursor: "pointer", fontSize: 11, fontFamily: "inherit"
            }}>🚪</button>
          </div>
          {/* Auto-logout toggle */}
          <button onClick={() => setAutoLogout(v => !v)} style={{
            width: "100%", marginTop: 6, padding: "7px", background: "transparent",
            border: `1px solid ${autoLogout ? "rgba(251,191,36,0.5)" : "rgba(255,255,255,0.15)"}`,
            color: autoLogout ? "#fbbf24" : "rgba(255,255,255,0.4)", borderRadius: 8,
            cursor: "pointer", fontSize: 11, fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 6
          }}>
            <span>{autoLogout ? "🔒" : "🔓"}</span>
            <span>Auto-logout: <b>{autoLogout ? "ON (60 min)" : "OFF"}</b></span>
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleUpload} style={{ display: "none" }} />
        </div>
      </div>

      {/* ── Main Content ── */}
      <div className="cm-main" style={{ marginLeft: "var(--sidebar-offset, 240px)", flex: 1, display: "flex", flexDirection: "column", minHeight: "100vh", background: "#f1f5f9" }}>

        {/* Top bar */}
        <div style={{ background: "#fff", borderBottom: "1px solid #e2e8f0", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
            <button className="cm-hamburger" onClick={() => setSidebarOpen(true)} style={{ color: "#0f2d5a" }}>
              <svg width="18" height="18" viewBox="0 0 18 18" fill="none"><rect y="2" width="18" height="2" rx="1" fill="currentColor"/><rect y="8" width="18" height="2" rx="1" fill="currentColor"/><rect y="14" width="18" height="2" rx="1" fill="currentColor"/></svg>
            </button>
            <div style={{ minWidth: 0 }}>
              <div className="cm-topbar-title" style={{ fontSize: 18, fontWeight: 800, color: "#0f2d5a", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                {page === "cost"     && "📊 BOM"}
                {page === "raw"      && "🧪 Raw Materials"}
                {page === "salesamt" && "💰 Sales Amount"}
                {page === "fg"       && "🏷️ Finished Goods"}
                {page === "balance"  && "📦 Item Balance"}
                {page === "rawcost"  && "💎 Raw Material Cost"}
                {page === "cashflow" && "💵 Cash Flow"}
                {page === "pricelist"&& "🏷️ Price List Type"}
                {page === "salesan"  && "📈 Sales Analysis"}
                {page === "pricelab" && "🔬 Price Lab"}
                {page === "coverage"  && "📊 Customer Coverage"}
              </div>
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 1 }}>BOM Cost + 35% Production Overhead</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {uploadMsg && <div style={{ background: uploadMsg.startsWith("✅") ? "#15803d" : "#dc2626", color: "#fff", padding: "5px 12px", borderRadius: 6, fontSize: 12 }}>{uploadMsg}</div>}
            {refreshMsg && (
              <div style={{ background: refreshMsg.startsWith("✅") ? "#f0fdf4" : "#fef2f2", border: `1px solid ${refreshMsg.startsWith("✅") ? "#86efac" : "#fca5a5"}`, color: refreshMsg.startsWith("✅") ? "#15803d" : "#dc2626", padding: "6px 12px", borderRadius: 8, fontSize: 12, display: "flex", alignItems: "center", gap: 8 }}>
                {refreshMsg.startsWith("✅") ? "✅" : "❌"} {refreshMsg.replace(/^✅|^❌/, "").trim()}
                <button onClick={() => setRefreshMsg("")} style={{ background: "none", border: "none", cursor: "pointer", color: "inherit", fontSize: 14, padding: 0 }}>✕</button>
              </div>
            )}
          </div>
        </div>

        {/* Auto-logout countdown warning */}
        {countdown !== null && (
          <div style={{ margin: "10px 20px 0", background: "#fef3c7", border: "1px solid #f59e0b", borderRadius: 10, padding: "10px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 18 }}>⏰</span>
              <span style={{ fontWeight: 700, color: "#92400e", fontSize: 13 }}>
                Logging out in <span style={{ fontFamily: "monospace", fontSize: 16, color: countdown <= 10 ? "#dc2626" : "#b45309" }}>{countdown}s</span> due to inactivity
              </span>
            </div>
            <button onClick={resetTimer} style={{ background: "#f59e0b", color: "#fff", border: "none", padding: "6px 16px", borderRadius: 7, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700 }}>
              Stay Logged In
            </button>
          </div>
        )}

        {/* Test API panel */}
        {testData && (
          <div style={{ margin: "12px 20px 0", background: "#fff", borderRadius: 12, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
            <div style={{ padding: "10px 16px", background: testData.ok ? "#f0fdf4" : "#fef2f2", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${testData.ok ? "#15803d" : "#dc2626"}` }}>
              <div style={{ fontWeight: 700, color: testData.ok ? "#15803d" : "#dc2626", fontSize: 13 }}>
                {testData.ok ? `✅ API Response — HTTP ${testData.status}` : `❌ API Error — ${testData.error}`}
              </div>
              <button onClick={() => setTestData(null)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 18, color: "#94a3b8" }}>✕</button>
            </div>
            {testData.ok && (
              <div style={{ padding: "12px 16px", overflowX: "auto", maxHeight: 400, overflowY: "auto" }}>
                <pre style={{ margin: 0, fontSize: 11, fontFamily: "monospace", color: "#1e293b", whiteSpace: "pre-wrap", wordBreak: "break-all" }}>
                  {JSON.stringify(testData.data, null, 2)}
                </pre>
              </div>
            )}
          </div>
        )}

        {/* Page content */}
        <div style={{ flex: 1 }}>

      {/* Raw Materials Page */}
      {page === "raw" && (() => {
        const filtered = RAW_MATERIALS
          .filter(r => r.code.toLowerCase().includes(rmSearch.toLowerCase()) || r.desc.includes(rmSearch))
          .sort((a, b) => {
            const av = rmSort === "code" ? a.code.toLowerCase() : rmSort === "lastCost" ? a.lastCost : a.avgCost;
            const bv = rmSort === "code" ? b.code.toLowerCase() : rmSort === "lastCost" ? b.lastCost : b.avgCost;
            return rmDir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
          });
        const rmSortBtn = (col, label) => (
          <button onClick={() => { setRmSort(col); setRmDir(d => d === "asc" && rmSort === col ? "desc" : "asc"); }}
            style={{ background: "none", border: "none", cursor: "pointer", fontWeight: rmSort === col ? 700 : 500, color: rmSort === col ? "#1d4ed8" : "#374151", fontSize: 12, fontFamily: "inherit" }}>
            {label}{rmSort === col ? (rmDir === "asc" ? " ↑" : " ↓") : ""}
          </button>
        );
        const overrideCount = Object.keys(costOverrides).length;
        return (
          <div style={{ padding: "16px 20px", display: "grid", gridTemplateColumns: selectedRaw ? "1fr 380px" : "1fr", gap: 14, alignItems: "start" }}>
            <div>{/* left column */}
            <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
              <div style={{ padding: "12px 16px", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 8, alignItems: "center" }}>
                <input value={rmSearch} onChange={e => setRmSearch(e.target.value)} placeholder="🔍  Search by item code..."
                  style={{ border: "1px solid #cbd5e1", borderRadius: 7, padding: "6px 12px", fontSize: 13, fontFamily: "inherit", flex: 1, minWidth: 200 }} />
                {overrideCount > 0 && (
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <span style={{ background: "#fef3c7", color: "#92400e", padding: "4px 10px", borderRadius: 20, fontSize: 12, fontWeight: 700 }}>
                      ✏️ {overrideCount} override{overrideCount > 1 ? "s" : ""} active
                    </span>
                    <button onClick={() => setCostOverrides({})}
                      style={{ background: "#fef2f2", border: "1px solid #fca5a5", color: "#dc2626", padding: "4px 10px", borderRadius: 7, cursor: "pointer", fontSize: 12, fontFamily: "inherit", fontWeight: 600 }}>
                      🔄 Reset All
                    </button>
                  </div>
                )}
                <span style={{ color: "#94a3b8", fontSize: 12 }}>{filtered.length} items</span>
              </div>
              <div style={{ padding: "6px 16px", background: "#f0fdf4", borderBottom: "1px solid #bbf7d0", fontSize: 12, color: "#15803d" }}>
                🧪 <b>Raw Materials</b> — ItemType = <b>R</b>, excluding MB-* • Click ✏️ to edit Last Cost — updates Cost Management instantly
              </div>
              <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 280px)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "9px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{rmSortBtn("code", "Item Code")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Description</th>
                      <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Type</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{rmSortBtn("lastCost", "Last Cost")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{rmSortBtn("avgCost", "Average Cost")}</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Diff</th>
                      <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>New Last Cost</th>
                      <th style={{ padding: "9px 8px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, idx) => {
                      const overridden = costOverrides[r.code] !== undefined;
                      const effectiveCost = overridden ? costOverrides[r.code] : r.lastCost;
                      const diff = r.avgCost - r.lastCost;
                      const diffPct = r.lastCost > 0 ? (diff / r.lastCost) * 100 : 0;
                      const diffColor = diff > 0.01 ? "#dc2626" : diff < -0.01 ? "#15803d" : "#64748b";
                      const isEditing = editingCode === r.code;
                      const isRawSelected = selectedRaw === r.code;
                      return (
                        <tr key={r.code} onClick={() => { if (!isEditing) setSelectedRaw(isRawSelected ? null : r.code); }}
                          style={{ borderBottom: "1px solid #f1f5f9", background: isRawSelected ? "#eff6ff" : overridden ? "#fefce8" : idx % 2 === 1 ? "#f8fafc" : "#fff", cursor: "pointer" }}>
                          <td style={{ padding: "8px 12px", fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 12 }}>
                            {r.code}
                            {overridden && <span style={{ marginLeft: 6, fontSize: 10, background: "#fbbf24", color: "#78350f", padding: "1px 6px", borderRadius: 10 }}>edited</span>}
                          </td>
                          <td style={{ padding: "8px 12px", textAlign: "right", direction: "rtl", color: "#374151", fontSize: 12 }}>{r.desc || <span style={{ color: "#94a3b8" }}>—</span>}</td>
                          <td style={{ padding: "8px", textAlign: "center" }}>
                            <span style={{ background: "#f0fdf4", color: "#15803d", padding: "2px 8px", borderRadius: 20, fontWeight: 700, fontSize: 11 }}>R</span>
                          </td>
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", color: overridden ? "#92400e" : "#374151", textDecoration: overridden ? "line-through" : "none" }}>{fmt(r.lastCost)}</td>
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#0f2d5a" }}>{fmt(r.avgCost)}</td>
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", color: diffColor, fontWeight: 600 }}>
                            {diff === 0 ? <span style={{ color: "#94a3b8" }}>—</span> : `${diff > 0 ? "+" : ""}${fmt(diff)} (${diffPct > 0 ? "+" : ""}${fmtPct(diffPct)}%)`}
                          </td>
                          <td style={{ padding: "6px 12px", textAlign: "center" }}>
                            {isEditing ? (
                              <input autoFocus type="number" step="0.01" value={editValue}
                                onChange={e => setEditValue(e.target.value)}
                                onKeyDown={e => { if (e.key === "Enter") handleSaveOverride(r.code); if (e.key === "Escape") { setEditingCode(null); setEditValue(""); } }}
                                style={{ width: 90, padding: "4px 8px", border: "2px solid #1d4ed8", borderRadius: 6, fontFamily: "monospace", fontSize: 13, color: "#1d4ed8", fontWeight: 700, textAlign: "right" }} />
                            ) : overridden ? (
                              <span style={{ fontFamily: "monospace", fontWeight: 800, color: "#b45309", fontSize: 13 }}>{fmt(effectiveCost)}</span>
                            ) : (
                              <span style={{ color: "#94a3b8", fontSize: 12 }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "6px 8px", textAlign: "center" }}>
                            <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                              {isEditing ? (
                                <>
                                  <button onClick={() => handleSaveOverride(r.code)}
                                    style={{ background: "#15803d", border: "none", color: "#fff", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>✓</button>
                                  <button onClick={() => { setEditingCode(null); setEditValue(""); }}
                                    style={{ background: "#f1f5f9", border: "1px solid #cbd5e1", color: "#64748b", padding: "4px 8px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>✕</button>
                                </>
                              ) : (
                                <>
                                  <button onClick={() => { setEditingCode(r.code); setEditValue(overridden ? costOverrides[r.code] : r.lastCost); }}
                                    style={{ background: "#eff6ff", border: "1px solid #bfdbfe", color: "#1d4ed8", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>✏️</button>
                                  {overridden && (
                                    <button onClick={() => handleResetOverride(r.code)}
                                      style={{ background: "#fef2f2", border: "1px solid #fca5a5", color: "#dc2626", padding: "4px 8px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>✕</button>
                                  )}
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Impact Panel */}
            {affectedItems.length > 0 && (
              <div style={{ marginTop: 14, background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
                <div style={{ padding: "12px 16px", background: "linear-gradient(135deg, #7c3aed, #a855f7)", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontWeight: 800, fontSize: 14 }}>📊 Impact Summary — {affectedItems.length} affected item{affectedItems.length > 1 ? "s" : ""}</div>
                    <div style={{ fontSize: 11, opacity: 0.8, marginTop: 2 }}>Before → After comparison based on current overrides</div>
                  </div>
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead style={{ background: "#f8fafc" }}>
                      <tr>
                        <th style={{ padding: "8px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Item Code</th>
                        <th style={{ padding: "8px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569", maxWidth: 160 }}>Description</th>
                        <th style={{ padding: "8px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Changed Material</th>
                        <th style={{ padding: "8px 10px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Material Cost</th>
                        <th style={{ padding: "8px 10px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Total Cost +35%</th>
                        <th style={{ padding: "8px 10px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Profit</th>
                        <th style={{ padding: "8px 10px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Profit %</th>
                        <th style={{ padding: "8px 10px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Δ Profit %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {affectedItems.map((item, idx) => {
                        const delta = Math.round((item.newPct - item.oldPct) * 10) / 10;
                        const deltaColor = delta >= 0 ? "#15803d" : "#dc2626";
                        const deltaBg = delta >= 0 ? "#f0fdf4" : "#fef2f2";
                        return (
                          <tr key={item.code} style={{ borderBottom: "1px solid #f1f5f9", background: idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
                            <td style={{ padding: "8px 12px", fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>{item.code}</td>
                            <td style={{ padding: "8px 12px", color: "#374151", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", direction: "rtl", textAlign: "right" }}>{item.desc}</td>
                            <td style={{ padding: "8px 12px", textAlign: "center" }}>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 3, justifyContent: "center" }}>
                                {item.changedChildren.map(c => (
                                  <span key={c} style={{ background: "#fef3c7", color: "#92400e", padding: "1px 6px", borderRadius: 10, fontSize: 10, fontFamily: "monospace", fontWeight: 600 }}>{c}</span>
                                ))}
                              </div>
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "right", fontFamily: "monospace" }}>
                              <div style={{ color: "#94a3b8", fontSize: 11 }}>{fmt(item.oldRaw)}</div>
                              <div style={{ fontWeight: 700, color: item.newRaw > item.oldRaw ? "#dc2626" : "#15803d" }}>→ {fmt(item.newRaw)}</div>
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "right", fontFamily: "monospace" }}>
                              <div style={{ color: "#94a3b8", fontSize: 11 }}>{fmt(item.oldBox)}</div>
                              <div style={{ fontWeight: 700, color: item.newBox > item.oldBox ? "#dc2626" : "#15803d" }}>→ {fmt(item.newBox)}</div>
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "right", fontFamily: "monospace" }}>
                              <div style={{ color: "#94a3b8", fontSize: 11 }}>{fmt(item.oldProfit)}</div>
                              <div style={{ fontWeight: 700, color: item.newProfit >= item.oldProfit ? "#15803d" : "#dc2626" }}>→ {fmt(item.newProfit)}</div>
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "center" }}>
                              <div style={{ color: "#94a3b8", fontSize: 11 }}>{fmtPct(item.oldPct)}%</div>
                              <div style={{ fontWeight: 700, color: profitColor(item.newPct) }}>→ {fmtPct(item.newPct)}%</div>
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "center" }}>
                              <span style={{ background: deltaBg, color: deltaColor, padding: "3px 10px", borderRadius: 20, fontWeight: 800, fontSize: 12, fontFamily: "monospace" }}>
                                {delta >= 0 ? "+" : ""}{fmtPct(delta)}%
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            </div>

            {/* BOM Usage Panel */}
            {selectedRaw && (() => {
              const rawItem = RAW_MATERIALS.find(r => r.code === selectedRaw);
              const usedIn = recalcData.filter(item => item.hasBOM && item.bomLines.some(l => l.child === selectedRaw));
              return (
                <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden", position: "sticky", top: 16 }}>
                  <div style={{ padding: "14px 16px", background: "linear-gradient(135deg, #0f2d5a, #1d4ed8)", color: "#fff" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "start" }}>
                      <div>
                        <div style={{ fontWeight: 800, fontSize: 13, fontFamily: "monospace" }}>{selectedRaw}</div>
                        <div style={{ fontSize: 12, opacity: 0.8, marginTop: 2, direction: "rtl" }}>{rawItem?.desc}</div>
                        <div style={{ marginTop: 8, fontSize: 12, opacity: 0.9 }}>Used in <b>{usedIn.length}</b> finished item{usedIn.length !== 1 ? "s" : ""}</div>
                      </div>
                      <button onClick={() => setSelectedRaw(null)}
                        style={{ background: "rgba(255,255,255,0.2)", border: "none", color: "#fff", width: 26, height: 26, borderRadius: 6, cursor: "pointer", fontSize: 14 }}>✕</button>
                    </div>
                  </div>
                  <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 280px)" }}>
                    {usedIn.length === 0 ? (
                      <div style={{ padding: "40px 16px", textAlign: "center", color: "#94a3b8" }}>
                        <div style={{ fontSize: 32 }}>📭</div>
                        <div style={{ marginTop: 8, fontWeight: 600, color: "#374151" }}>Not used in any BOM</div>
                      </div>
                    ) : usedIn.map((item, idx) => {
                      const line = item.bomLines.find(l => l.child === selectedRaw);
                      const contribution = item.rawBoxCost > 0 ? (line.lineCost / item.rawBoxCost * 100) : 0;
                      return (
                        <div key={item.code} style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9", background: idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                            <span style={{ fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 12 }}>{item.code}</span>
                            <span style={{ background: profitBg(item.profitPct), color: profitColor(item.profitPct), padding: "1px 8px", borderRadius: 20, fontWeight: 700, fontSize: 11, fontFamily: "monospace" }}>{fmtPct(item.profitPct)}%</span>
                          </div>
                          <div style={{ fontSize: 11, color: "#64748b", direction: "rtl", textAlign: "right", marginBottom: 6 }}>{item.desc}</div>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 4 }}>
                            {[
                              { l: "Qty", v: fmtQty(line.qty) },
                              { l: "Line Cost", v: fmt(line.lineCost) },
                              { l: "% of Material", v: fmtPct(contribution) + "%" },
                            ].map((s, i) => (
                              <div key={i} style={{ background: "#f8fafc", borderRadius: 6, padding: "4px 6px", textAlign: "center" }}>
                                <div style={{ fontSize: 9, color: "#94a3b8" }}>{s.l}</div>
                                <div style={{ fontWeight: 700, fontSize: 11, fontFamily: "monospace", color: "#0f2d5a" }}>{s.v}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}
          </div>
        );
      })()}

      {/* Cost Management Page */}
      {page === "cost" && <>
      {/* Stats */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 10, padding: "14px 20px" }}>
        {[
          { label: "Total Items", value: stats.total, icon: "📦", color: "#1d4ed8", bg: "#eff6ff" },
          { label: "With BOM", value: stats.withBOM, icon: "✅", color: "#15803d", bg: "#f0fdf4" },
          { label: "No BOM", value: stats.noBOM, icon: "❌", color: "#dc2626", bg: "#fef2f2" },
          { label: "Partial BOM", value: stats.partial, icon: "⚠️", color: "#d97706", bg: "#fffbeb" },
          { label: "Avg Profit %", value: fmtPct(stats.avgProfit) + "%", icon: "📊", color: "#7c3aed", bg: "#f5f3ff" },
          { label: "Profit < 40%", value: stats.low, icon: "🔴", color: "#dc2626", bg: "#fef2f2" },
        ].map((s, i) => (
          <div key={i} style={{ background: s.bg, borderRadius: 10, padding: "12px 14px", border: `1px solid ${s.color}22` }}>
            <div style={{ fontSize: 20 }}>{s.icon}</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: s.color, lineHeight: 1.3 }}>{s.value}</div>
            <div style={{ fontSize: 11, color: "#64748b" }}>{s.label}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: selected ? "1fr 420px" : "1fr", gap: 14, padding: "0 20px 20px" }}>
        {/* Table */}
        <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
          {/* Filters */}
          <div style={{ padding: "12px 16px", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍  Search by code or name..."
              style={{ border: "1px solid #cbd5e1", borderRadius: 7, padding: "6px 12px", fontSize: 13, fontFamily: "inherit", flex: 1, minWidth: 180 }} />
            {["all", "with", "without", "partial"].map(f => (
              <button key={f} onClick={() => setFilter(f)}
                style={{ padding: "6px 12px", borderRadius: 7, border: "1px solid " + (filterBOM === f ? "#1d4ed8" : "#cbd5e1"), background: filterBOM === f ? "#1d4ed8" : "#fff", color: filterBOM === f ? "#fff" : "#374151", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600 }}>
                {{ all: "All", with: "✅ With BOM", without: "❌ No BOM", partial: "⚠️ Partial" }[f]}
              </button>
            ))}
            <div style={{ display: "flex", alignItems: "center", gap: 6, background: discount > 0 ? "#fef3c7" : "#f8fafc", border: "1px solid " + (discount > 0 ? "#fbbf24" : "#cbd5e1"), borderRadius: 7, padding: "4px 10px" }}>
              <span style={{ fontSize: 12, color: "#475569", whiteSpace: "nowrap" }}>🏷️ Discount</span>
              <input type="number" min="0" max="100" step="0.5" value={discount}
                onChange={e => setDiscount(Math.min(100, Math.max(0, parseFloat(e.target.value) || 0)))}
                style={{ width: 52, padding: "2px 6px", border: "1px solid #cbd5e1", borderRadius: 5, fontFamily: "monospace", fontSize: 13, fontWeight: 700, color: discount > 0 ? "#b45309" : "#374151", textAlign: "right" }} />
              <span style={{ fontSize: 12, color: "#475569" }}>%</span>
              {discount > 0 && <button onClick={() => setDiscount(0)} style={{ background: "none", border: "none", cursor: "pointer", color: "#94a3b8", fontSize: 12, padding: 0 }}>✕</button>}
            </div>
            <button onClick={() => {
              const wb = XLSX.utils.book_new();
              const rows = [
                ["#", "Item Code", "Item Description", "Selling Conversion", "Selling Price", "Material Cost", "Total Cost +35%", "Profit", "Profit %",
                  ...(discount > 0 ? [`Net Price (${fmtPct(discount)}% off)`, "Profit % after disc."] : []),
                  "BOM Status", "Missing Costs"
                ],
                ...filtered.map((item, idx) => {
                  const netPrice = discount > 0 && item.hasBOM ? Math.round(item.sp * (1 - discount / 100) * 100) / 100 : null;
                  const profitAfterDisc = netPrice !== null ? Math.round((netPrice - item.boxCost) * 100) / 100 : null;
                  const pctAfterDisc = profitAfterDisc !== null && netPrice > 0 ? Math.round((profitAfterDisc / netPrice * 100) * 10) / 10 : null;
                  return [
                    idx + 1,
                    item.code,
                    item.desc || "",
                    item.conv,
                    item.sp,
                    item.hasBOM ? item.rawBoxCost : "",
                    item.hasBOM ? item.boxCost : "",
                    item.hasBOM ? item.profit : "",
                    item.hasBOM ? item.profitPct : "",
                    ...(discount > 0 ? [netPrice ?? "", pctAfterDisc ?? ""] : []),
                    !item.hasBOM ? "No BOM" : item.missing.length > 0 ? "Partial" : "OK",
                    item.missing.join(", "),
                  ];
                })
              ];
              const ws = XLSX.utils.aoa_to_sheet(rows);
              ws["!cols"] = [4,14,24,8,10,12,12,10,8,...(discount>0?[10,10]:[]),8,20].map(w=>({wch:w}));
              XLSX.utils.book_append_sheet(wb, ws, "BOM");
              XLSX.writeFile(wb, "Cost_Management.xlsx");
            }}
              style={{ background: "#15803d", border: "none", color: "#fff", padding: "6px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit", whiteSpace: "nowrap" }}>
              ⬇️ Export Excel
            </button>
            <button onClick={async () => {
              // Load libs from CDN
              const loadScript = (src) => new Promise((res, rej) => {
                if (document.querySelector(`script[src="${src}"]`)) return res();
                const s = document.createElement("script");
                s.src = src; s.onload = res; s.onerror = rej;
                document.head.appendChild(s);
              });
              await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
              await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");

              const { jsPDF } = window.jspdf;

              // Build a clean off-screen HTML table with all data
              const container = document.createElement("div");
              container.style.cssText = "position:fixed;left:-9999px;top:0;background:#fff;padding:20px;font-family:Arial,sans-serif;width:1100px;direction:ltr";

              const today = new Date().toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" });
              container.innerHTML = `
                <div style="background:#1e293b;color:#fff;padding:10px 16px;border-radius:6px 6px 0 0;display:flex;justify-content:space-between;align-items:center">
                  <span style="font-size:16px;font-weight:700">📊 BOM Cost Report</span>
                  <span style="font-size:12px;opacity:0.8">${today} &nbsp;·&nbsp; ${filtered.length} items${discount > 0 ? ` &nbsp;·&nbsp; Discount: ${fmtPct(discount)}%` : ""}</span>
                </div>
                <div style="background:#f8fafc;border:1px solid #e2e8f0;border-top:none;padding:8px 16px;display:flex;gap:24px;font-size:12px;color:#475569">
                  <span>Total: <b>${stats.total}</b></span>
                  <span>✅ With BOM: <b>${stats.withBOM}</b></span>
                  <span>❌ No BOM: <b>${stats.noBOM}</b></span>
                  <span>⚠️ Partial: <b>${stats.partial}</b></span>
                  <span>Avg Profit: <b>${fmtPct(stats.avgProfit)}%</b></span>
                  <span>Low Profit: <b style="color:#dc2626">${stats.low}</b></span>
                </div>
                <table style="width:100%;border-collapse:collapse;font-size:11px;margin-top:0">
                  <thead>
                    <tr style="background:#7c3aed;color:#fff">
                      <th style="padding:7px 8px;text-align:center;width:30px">#</th>
                      <th style="padding:7px 8px;text-align:left;width:110px">Code</th>
                      <th style="padding:7px 8px;text-align:left">Description</th>
                      <th style="padding:7px 8px;text-align:center;width:50px">Conv.</th>
                      <th style="padding:7px 8px;text-align:right;width:75px">Sell Price</th>
                      <th style="padding:7px 8px;text-align:right;width:85px">Material Cost</th>
                      <th style="padding:7px 8px;text-align:right;width:85px">Total +35%</th>
                      <th style="padding:7px 8px;text-align:right;width:75px">Profit</th>
                      <th style="padding:7px 8px;text-align:center;width:65px">Profit %</th>
                      ${discount > 0 ? `<th style="padding:7px 8px;text-align:right;width:85px">Net Price</th><th style="padding:7px 8px;text-align:center;width:75px">% After Disc</th>` : ""}
                      <th style="padding:7px 8px;text-align:center;width:60px">BOM</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${filtered.map((item, idx) => {
                      const netPrice = discount > 0 && item.hasBOM ? Math.round(item.sp * (1 - discount/100)*100)/100 : null;
                      const profitAfterDisc = netPrice !== null ? Math.round((netPrice - item.boxCost)*100)/100 : null;
                      const pctAfterDisc = profitAfterDisc !== null && netPrice > 0 ? Math.round((profitAfterDisc/netPrice*100)*10)/10 : null;
                      const bomStatus = !item.hasBOM ? "No BOM" : item.missing.length > 0 ? "Partial" : "OK";
                      const bomColor = bomStatus === "OK" ? "#15803d" : bomStatus === "Partial" ? "#b45309" : "#dc2626";
                      const pct = item.hasBOM ? item.profitPct : null;
                      const pctColor = pct === null ? "#94a3b8" : pct >= 40 ? "#15803d" : pct >= 20 ? "#b45309" : "#dc2626";
                      const rowBg = !item.hasBOM ? "#fef2f2" : idx % 2 === 0 ? "#fff" : "#f8fafc";
                      return `<tr style="background:${rowBg}">
                        <td style="padding:6px 8px;text-align:center;color:#94a3b8">${idx+1}</td>
                        <td style="padding:6px 8px;font-family:monospace;font-weight:700;color:#0f2d5a">${item.code}</td>
                        <td style="padding:6px 8px;color:#1e293b">${item.desc || ""}</td>
                        <td style="padding:6px 8px;text-align:center">${item.conv}</td>
                        <td style="padding:6px 8px;text-align:right">${item.hasBOM ? item.sp : ""}</td>
                        <td style="padding:6px 8px;text-align:right">${item.hasBOM ? item.rawBoxCost : ""}</td>
                        <td style="padding:6px 8px;text-align:right">${item.hasBOM ? item.boxCost : ""}</td>
                        <td style="padding:6px 8px;text-align:right">${item.hasBOM ? item.profit : ""}</td>
                        <td style="padding:6px 8px;text-align:center;font-weight:700;color:${pctColor}">${pct !== null ? pct + "%" : "—"}</td>
                        ${discount > 0 ? `<td style="padding:6px 8px;text-align:right">${netPrice ?? ""}</td><td style="padding:6px 8px;text-align:center;font-weight:700;color:${pctColor}">${pctAfterDisc != null ? pctAfterDisc+"%" : ""}</td>` : ""}
                        <td style="padding:6px 8px;text-align:center;font-weight:700;color:${bomColor}">${bomStatus}</td>
                      </tr>`;
                    }).join("")}
                  </tbody>
                </table>
                <div style="margin-top:8px;font-size:10px;color:#94a3b8;text-align:center">
                  Sila System — Cost Management &nbsp;·&nbsp; Total Cost = Material Cost × 1.35 (includes 35% production overhead)
                </div>
              `;
              document.body.appendChild(container);

              try {
                const canvas = await window.html2canvas(container, { scale: 2, useCORS: true, backgroundColor: "#fff" });
                const imgData = canvas.toDataURL("image/png");
                const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
                const pageW = pdf.internal.pageSize.getWidth();
                const pageH = pdf.internal.pageSize.getHeight();
                const margin = 8;
                const usableW = pageW - margin * 2;
                const totalH = (canvas.height / canvas.width) * usableW;
                const pagesNeeded = Math.ceil(totalH / (pageH - margin * 2));

                for (let i = 0; i < pagesNeeded; i++) {
                  if (i > 0) pdf.addPage();
                  const srcY = i * (canvas.height / pagesNeeded);
                  const srcH = canvas.height / pagesNeeded;
                  const pageCanvas = document.createElement("canvas");
                  pageCanvas.width = canvas.width;
                  pageCanvas.height = srcH;
                  const ctx = pageCanvas.getContext("2d");
                  ctx.drawImage(canvas, 0, -srcY);
                  const pageImg = pageCanvas.toDataURL("image/png");
                  const sliceH = (srcH / canvas.width) * usableW;
                  pdf.addImage(pageImg, "PNG", margin, margin, usableW, sliceH);
                  pdf.setFontSize(7);
                  pdf.setTextColor(148, 163, 184);
                  pdf.text(`Page ${i+1} of ${pagesNeeded}`, pageW/2, pageH-3, { align:"center" });
                }

                pdf.save(`BOM_Report_${new Date().toISOString().slice(0,10)}.pdf`);
              } finally {
                document.body.removeChild(container);
              }
            }}
              style={{ background: "#dc2626", border: "none", color: "#fff", padding: "6px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit", whiteSpace: "nowrap" }}>
              🖨️ Export PDF
            </button>
            <span style={{ color: "#94a3b8", fontSize: 12 }}>{filtered.length} items</span>
          </div>

          {/* Cost formula note */}
          <div style={{ padding: "6px 16px", background: "#eff6ff", borderBottom: "1px solid #bfdbfe", fontSize: 12, color: "#1e40af" }}>
            💡 <b>Total Cost</b> = BOM Material Cost × <b>1.35</b> (includes 35% production overhead)
          </div>

          <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 330px)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 10 }}>
                <tr>
                  <th style={{ padding: "9px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("code", "#  Code")}</th>
                  <th style={{ padding: "9px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Description</th>
                  <th style={{ padding: "9px 8px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Conv.</th>
                  <th style={{ padding: "9px 8px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("sp", "Sell Price")}</th>
                  <th style={{ padding: "9px 8px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("rawBoxCost", "Material Cost")}</th>
                  <th style={{ padding: "9px 8px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("boxCost", "Total Cost +35%")}</th>
                  <th style={{ padding: "9px 8px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("profit", "Profit")}</th>
                  <th style={{ padding: "9px 8px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>{sortBtn("profitPct", "Profit %")}</th>
                  {discount > 0 && <th style={{ padding: "9px 8px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#b45309", background: "#fffbeb", whiteSpace: "nowrap" }}>Net Price ({fmtPct(discount)}% off)</th>}
                  {discount > 0 && <th style={{ padding: "9px 8px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#b45309", background: "#fffbeb", whiteSpace: "nowrap" }}>Profit % after disc.</th>}
                  <th style={{ padding: "9px 8px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>BOM</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((item, idx) => {
                  const isSelected = selected === item.code;
                  const isAffected = affectedSet.has(item.code);
                  const aff = isAffected ? affectedMap[item.code] : null;
                  const rowBg = !item.hasBOM ? "#fef2f2" : isAffected ? "#fefce8" : isSelected ? "#eff6ff" : idx % 2 === 1 ? "#f8fafc" : "#fff";
                  const cell = (oldVal, newVal, isNum = true) => isAffected ? (
                    <div style={{ lineHeight: 1.3 }}>
                      <div style={{ color: "#94a3b8", fontSize: 11, textDecoration: "line-through" }}>{oldVal}</div>
                      <div style={{ fontWeight: 700, color: (isNum ? newVal > oldVal : false) ? "#dc2626" : "#15803d" }}>{newVal}</div>
                    </div>
                  ) : oldVal;
                  return (
                    <tr key={item.code} onClick={() => { setSelected(isSelected ? null : item.code); setScenario({}); setTab("bom"); }}
                      style={{ background: rowBg, cursor: "pointer", borderBottom: "1px solid #f1f5f9" }}>
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 12 }}>
                        {item.code}
                        {isAffected && <span style={{ marginLeft: 5, fontSize: 9, background: "#fbbf24", color: "#78350f", padding: "1px 5px", borderRadius: 8 }}>✏️</span>}
                      </td>
                      <td style={{ padding: "8px 12px", color: "#374151", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", direction: "rtl", textAlign: "right" }}>{item.desc || <span style={{ color: "#94a3b8" }}>—</span>}</td>
                      <td style={{ padding: "8px", textAlign: "center", color: "#64748b", fontFamily: "monospace" }}>{item.conv}</td>
                      <td style={{ padding: "8px", textAlign: "right", fontFamily: "monospace", color: "#0f2d5a" }}>{fmt(item.sp)}</td>
                      <td style={{ padding: "8px", textAlign: "right", fontFamily: "monospace", color: "#475569" }}>
                        {item.hasBOM ? cell(fmt(aff?.oldRaw ?? item.rawBoxCost), fmt(aff?.newRaw ?? item.rawBoxCost), true) : <span style={{ color: "#94a3b8" }}>—</span>}
                      </td>
                      <td style={{ padding: "8px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: item.hasBOM ? "#0f2d5a" : "#94a3b8" }}>
                        {item.hasBOM ? cell(fmt(aff?.oldBox ?? item.boxCost), fmt(aff?.newBox ?? item.boxCost), true) : "—"}
                      </td>
                      <td style={{ padding: "8px", textAlign: "right", fontFamily: "monospace" }}>
                        {item.hasBOM ? cell(fmt(aff?.oldProfit ?? item.profit), fmt(aff?.newProfit ?? item.profit), false) : <span style={{ color: "#94a3b8" }}>—</span>}
                      </td>
                      <td style={{ padding: "8px", textAlign: "center" }}>
                        {item.hasBOM ? (
                          isAffected ? (
                            <div style={{ lineHeight: 1.3 }}>
                              <div style={{ color: "#94a3b8", fontSize: 11, textDecoration: "line-through", fontFamily: "monospace" }}>{fmtPct(aff.oldPct)}%</div>
                              <span style={{ background: profitBg(aff.newPct), color: profitColor(aff.newPct), padding: "2px 9px", borderRadius: 20, fontWeight: 700, fontSize: 12, fontFamily: "monospace" }}>{fmtPct(aff.newPct)}%</span>
                            </div>
                          ) : (
                            <span style={{ background: profitBg(item.profitPct), color: profitColor(item.profitPct), padding: "2px 9px", borderRadius: 20, fontWeight: 700, fontSize: 12, fontFamily: "monospace" }}>{fmtPct(item.profitPct)}%</span>
                          )
                        ) : <span style={{ color: "#94a3b8" }}>—</span>}
                      </td>
                      {discount > 0 && !item.hasBOM && <td style={{ padding: "8px", background: "#fffbeb", textAlign: "right", color: "#94a3b8" }}>—</td>}
                      {discount > 0 && !item.hasBOM && <td style={{ padding: "8px", background: "#fffbeb", textAlign: "center", color: "#94a3b8" }}>—</td>}
                      {discount > 0 && item.hasBOM && (() => {
                        const netPrice = Math.round(item.sp * (1 - discount / 100) * 100) / 100;
                        const profitAfterDisc = Math.round((netPrice - item.boxCost) * 100) / 100;
                        const pctAfterDisc = netPrice > 0 ? Math.round((profitAfterDisc / netPrice * 100) * 10) / 10 : 0;
                        return [
                          <td key="np" style={{ padding: "8px", textAlign: "right", fontFamily: "monospace", background: "#fffbeb", color: "#0f2d5a", fontWeight: 600 }}>{fmt(netPrice)}</td>,
                          <td key="pp" style={{ padding: "8px", textAlign: "center", background: "#fffbeb" }}>
                            <div style={{ lineHeight: 1.3 }}>
                              <span style={{ background: profitBg(pctAfterDisc), color: profitColor(pctAfterDisc), padding: "2px 8px", borderRadius: 20, fontWeight: 700, fontSize: 12, fontFamily: "monospace" }}>{fmtPct(pctAfterDisc)}%</span>
                              <div style={{ fontSize: 10, color: "#b45309", marginTop: 1 }}>{fmtPct(pctAfterDisc - item.profitPct)}pp</div>
                            </div>
                          </td>
                        ];
                      })()}
                      <td style={{ padding: "8px", textAlign: "center" }}>
                        {!item.hasBOM ? "❌" : item.missing.length > 0 ? <span title={item.missing.join(", ")} style={{ cursor: "help" }}>⚠️</span> : "✅"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* Detail panel */}
        {selItem && (
          <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", display: "flex", flexDirection: "column", maxHeight: "calc(100vh - 240px)", overflow: "hidden" }}>
            {/* Header */}
            <div style={{ padding: "14px 18px", background: "linear-gradient(135deg, #0f2d5a, #1d4ed8)", color: "#fff" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "start" }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15, fontFamily: "monospace" }}>{selItem.code}</div>
                  <div style={{ fontSize: 12, opacity: 0.8, marginTop: 2, direction: "rtl" }}>{selItem.desc}</div>
                </div>
                <button onClick={() => setSelected(null)} style={{ background: "rgba(255,255,255,0.2)", border: "none", color: "#fff", width: 26, height: 26, borderRadius: 6, cursor: "pointer", fontSize: 14 }}>✕</button>
              </div>
              {selItem.hasBOM && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6, marginTop: 10 }}>
                  {[
                    { l: "Sell Price", v: fmt(selItem.sp) },
                    { l: "Material", v: fmt(selItem.rawBoxCost) },
                    { l: "+35% Total", v: fmt(selItem.boxCost) },
                    { l: "Profit " + fmtPct(selItem.profitPct) + "%", v: fmt(selItem.profit) },
                  ].map((s, i) => (
                    <div key={i} style={{ background: "rgba(255,255,255,0.15)", borderRadius: 7, padding: "7px 8px", textAlign: "center" }}>
                      <div style={{ fontSize: 10, opacity: 0.8 }}>{s.l}</div>
                      <div style={{ fontWeight: 700, fontSize: 13, marginTop: 1, fontFamily: "monospace" }}>{s.v}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {selItem.hasBOM && (
              <div style={{ display: "flex", borderBottom: "1px solid #e2e8f0" }}>
                {[["bom", "📋 BOM Breakdown"], ["whatif", "🧮 What-If Scenario"]].map(([t, l]) => (
                  <button key={t} onClick={() => setTab(t)}
                    style={{ flex: 1, padding: "10px", border: "none", borderBottom: `3px solid ${tab === t ? "#1d4ed8" : "transparent"}`, background: "none", cursor: "pointer", color: tab === t ? "#1d4ed8" : "#64748b", fontWeight: tab === t ? 700 : 500, fontSize: 12, fontFamily: "inherit" }}>
                    {l}
                  </button>
                ))}
              </div>
            )}

            <div style={{ overflowY: "auto", flex: 1, padding: "14px 16px" }}>
              {!selItem.hasBOM ? (
                <div style={{ textAlign: "center", padding: "40px 16px", color: "#94a3b8" }}>
                  <div style={{ fontSize: 40 }}>📭</div>
                  <div style={{ fontWeight: 600, color: "#374151", marginTop: 8 }}>No Bill of Materials</div>
                  <div style={{ fontSize: 12, marginTop: 4 }}>This item has no BOM entry</div>
                </div>
              ) : tab === "bom" ? (
                <>
                  {selItem.missing.length > 0 && (
                    <div style={{ background: "#fffbeb", border: "1px solid #fcd34d", borderRadius: 8, padding: "8px 12px", marginBottom: 12, fontSize: 12 }}>
                      <b style={{ color: "#92400e" }}>⚠️ Missing costs:</b> <span style={{ color: "#78350f" }}>{selItem.missing.join(" • ")}</span>
                    </div>
                  )}
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: "#f8fafc" }}>
                        <th style={{ padding: "7px 8px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Child Item</th>
                        <th style={{ padding: "7px 6px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Qty</th>
                        <th style={{ padding: "7px 6px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Unit Cost</th>
                        <th style={{ padding: "7px 6px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Line Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {selItem.bomLines.map((l, i) => (
                        <tr key={i} style={{ borderBottom: "1px solid #f1f5f9", background: l.missing ? "#fef2f2" : "transparent" }}>
                          <td style={{ padding: "6px 8px", fontWeight: 600, color: l.missing ? "#dc2626" : "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>{l.child}{l.missing && " ⚠️"}</td>
                          <td style={{ padding: "6px", textAlign: "right", color: "#64748b", fontFamily: "monospace" }}>{fmtQty(l.qty)}</td>
                          <td style={{ padding: "6px", textAlign: "right", fontFamily: "monospace" }}>{l.unitCost !== null ? fmt(l.unitCost) : <span style={{ color: "#dc2626" }}>—</span>}</td>
                          <td style={{ padding: "6px", textAlign: "right", fontFamily: "monospace", fontWeight: 600 }}>{fmt(l.lineCost)}</td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{ background: "#f8fafc", borderTop: "2px solid #e2e8f0" }}>
                        <td colSpan={3} style={{ padding: "8px", fontWeight: 700 }}>Material Cost (Box)</td>
                        <td style={{ padding: "8px 6px", fontWeight: 800, color: "#1d4ed8", fontFamily: "monospace", textAlign: "right" }}>{fmt(selItem.rawBoxCost)}</td>
                      </tr>
                      <tr style={{ background: "#eff6ff" }}>
                        <td colSpan={3} style={{ padding: "8px", fontWeight: 700, color: "#1e40af" }}>+ 35% Production</td>
                        <td style={{ padding: "8px 6px", fontWeight: 800, color: "#1e40af", fontFamily: "monospace", textAlign: "right" }}>{fmt(selItem.rawBoxCost * 0.35)}</td>
                      </tr>
                      <tr style={{ background: "#dbeafe" }}>
                        <td colSpan={3} style={{ padding: "8px", fontWeight: 700, color: "#1d4ed8" }}>Total Cost (Box)</td>
                        <td style={{ padding: "8px 6px", fontWeight: 800, color: "#1d4ed8", fontFamily: "monospace", textAlign: "right" }}>{fmt(selItem.boxCost)}</td>
                      </tr>
                    </tfoot>
                  </table>
                </>
              ) : (
                <div>
                  <div style={{ background: "#eff6ff", borderRadius: 8, padding: "10px 12px", marginBottom: 14, fontSize: 12, color: "#1e40af" }}>
                    💡 Adjust material costs below to simulate impact on profit (production 35% is applied automatically)
                  </div>
                  {scenarioResult && (
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
                      {[
                        { l: "New Total Cost", v: fmt(scenarioResult.newBoxCost), d: fmt(selItem.boxCost), up: scenarioResult.newBoxCost > selItem.boxCost },
                        { l: "New Profit", v: fmt(scenarioResult.newProfit), d: fmt(selItem.profit), up: scenarioResult.newProfit > selItem.profit },
                        { l: "Profit %", v: fmtPct(scenarioResult.newPct) + "%", d: fmtPct(selItem.profitPct) + "%", up: scenarioResult.newPct > selItem.profitPct },
                      ].map((s, i) => (
                        <div key={i} style={{ background: "#f8fafc", borderRadius: 8, padding: "8px", border: "1px solid #e2e8f0", textAlign: "center" }}>
                          <div style={{ fontSize: 10, color: "#64748b" }}>{s.l}</div>
                          <div style={{ fontSize: 14, fontWeight: 800, color: "#0f2d5a", fontFamily: "monospace" }}>{s.v}</div>
                          <div style={{ fontSize: 10, color: "#94a3b8" }}>was {s.d} {s.up ? "↑" : "↓"}</div>
                        </div>
                      ))}
                    </div>
                  )}
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: "#f8fafc" }}>
                        <th style={{ padding: "7px 8px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Child Item</th>
                        <th style={{ padding: "7px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Qty</th>
                        <th style={{ padding: "7px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>New Unit Cost</th>
                      </tr>
                    </thead>
                    <tbody>
                      {selItem.bomLines.filter(l => !l.missing).map((l, i) => (
                        <tr key={i} style={{ borderBottom: "1px solid #f1f5f9" }}>
                          <td style={{ padding: "5px 8px", fontWeight: 600, color: "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>{l.child}</td>
                          <td style={{ padding: "5px", textAlign: "right", fontFamily: "monospace", color: "#64748b" }}>{fmtQty(l.qty)}</td>
                          <td style={{ padding: "4px 6px" }}>
                            <input type="number" step="0.01" defaultValue={l.unitCost}
                              onChange={e => setScenario(s => ({ ...s, [l.child]: e.target.value }))}
                              style={{ width: "100%", padding: "4px 7px", border: "1px solid " + (scenario[l.child] !== undefined ? "#1d4ed8" : "#cbd5e1"), borderRadius: 5, fontFamily: "monospace", fontSize: 12, color: scenario[l.child] !== undefined ? "#1d4ed8" : "#374151", fontWeight: scenario[l.child] !== undefined ? 700 : 400 }} />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <button onClick={() => setScenario({})} style={{ marginTop: 12, width: "100%", padding: "8px", background: "#f1f5f9", border: "1px solid #cbd5e1", borderRadius: 7, cursor: "pointer", fontFamily: "inherit", fontSize: 12 }}>🔄 Reset to Original</button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
      </>}

      {/* Sales Amount Page */}
      {page === "salesamt" && (() => {
        const costMap = Object.fromEntries(recalcData.filter(i => i.hasBOM).map(i => [i.code, i.rawBoxCost]));
        const convMap = Object.fromEntries(recalcData.map(i => [i.code, i.conv]));
        const descMap = Object.fromEntries(recalcData.map(i => [i.code, i.desc]));
        const salesSource = (window.__apiSalesData && window.__apiSalesData.length > 0) ? window.__apiSalesData : SALES_DATA;
        const rows = salesSource.map(r => {
          const boxCost = costMap[r.code] ?? costMap[r.code.trim()] ?? null;
          const conv = convMap[r.code] ?? convMap[r.code.trim()] ?? null;
          const unitCost = boxCost !== null && conv ? Math.round((boxCost / conv) * 100000) / 100000 : null;
          const desc = descMap[r.code] ?? descMap[r.code.trim()] ?? "";
          const totalCost = unitCost !== null ? Math.round(unitCost * r.qty * 100) / 100 : null;
          const grossProfit = totalCost !== null ? Math.round((r.amount - totalCost) * 100) / 100 : null;
          const profitPct = grossProfit !== null && r.amount !== 0 ? Math.round((grossProfit / r.amount * 100) * 10) / 10 : null;
          return { ...r, desc, boxCost, conv, unitCost, totalCost, grossProfit, profitPct };
        });
        const totalAmount = rows.reduce((s, r) => s + r.amount, 0);
        const totalCostSum = rows.filter(r => r.totalCost !== null).reduce((s, r) => s + r.totalCost, 0);
        const totalProfit = rows.filter(r => r.grossProfit !== null).reduce((s, r) => s + r.grossProfit, 0);
        const matchedRows = rows.filter(r => r.unitCost !== null);
        const matchedAmount = matchedRows.reduce((s, r) => s + r.amount, 0);
        return (
          <div style={{ padding: "16px 20px" }}>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 10, marginBottom: 14 }}>
              {[
                { label: "Total Items", value: rows.length, icon: "📦", color: "#1d4ed8", bg: "#eff6ff" },
                { label: "Total Sales Amount", value: fmt(totalAmount), icon: "💰", color: "#15803d", bg: "#f0fdf4" },
                { label: "Sales Amount (matched)", value: fmt(matchedAmount), icon: "🎯", color: "#0891b2", bg: "#ecfeff" },
                { label: "Total Cost (matched)", value: fmt(totalCostSum), icon: "🏭", color: "#7c3aed", bg: "#f5f3ff" },
                { label: "Gross Profit (matched)", value: fmt(totalProfit), icon: "📈", color: totalProfit >= 0 ? "#15803d" : "#dc2626", bg: totalProfit >= 0 ? "#f0fdf4" : "#fef2f2" },
              ].map((s, i) => (
                <div key={i} style={{ background: s.bg, borderRadius: 10, padding: "12px 14px", border: `1px solid ${s.color}22` }}>
                  <div style={{ fontSize: 20 }}>{s.icon}</div>
                  <div style={{ fontSize: 18, fontWeight: 800, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>{s.label}</div>
                </div>
              ))}
            </div>

            <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
              <div style={{ padding: "8px 16px", background: "#eff6ff", borderBottom: "1px solid #bfdbfe", fontSize: 12, color: "#1e40af", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>💡 <b>Total Cost</b> = Unit Cost (Material only) × Qty &nbsp;|&nbsp; {matchedRows.length} of {rows.length} items matched to BOM cost data</span>
                <button onClick={() => {
                  const wb = XLSX.utils.book_new();
                  const wsRows = [
                    ["#", "Item Code", "Item Description", "Qty", "Sales Amount", "Conv.", "Box Cost (Material)", "Unit Cost", "Total Cost", "Gross Profit", "Profit %"],
                    ...rows.map((r, idx) => [
                      idx + 1,
                      r.code,
                      r.desc || "",
                      r.qty,
                      r.amount,
                      r.conv ?? "",
                      r.boxCost ?? "",
                      r.unitCost ?? "",
                      r.totalCost ?? "",
                      r.grossProfit ?? "",
                      r.profitPct ?? "",
                    ]),
                    [],
                    ["", "TOTAL", "", rows.reduce((s,r)=>s+r.qty,0), totalAmount, "", "", "", totalCostSum, totalProfit, totalAmount > 0 ? Math.round(totalProfit/totalAmount*100*10)/10 : ""],
                  ];
                  const ws = XLSX.utils.aoa_to_sheet(wsRows);
                  ws["!cols"] = [4,14,24,8,12,6,14,12,12,12,8].map(w=>({wch:w}));
                  XLSX.utils.book_append_sheet(wb, ws, "Sales Amount");
                  XLSX.writeFile(wb, "Sales_Amount.xlsx");
                }}
                  style={{ background: "#15803d", border: "none", color: "#fff", padding: "5px 12px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit", whiteSpace: "nowrap" }}>
                  ⬇️ Export Excel
                </button>
              </div>
              <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 320px)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "9px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>#</th>
                      <th style={{ padding: "9px 12px", textAlign: "left", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Item Code</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Description</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Qty</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Sales Amount</th>
                      <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#64748b" }}>Conv.</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#1d4ed8" }}>Box Cost (Material)</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#7c3aed" }}>Unit Cost</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#7c3aed" }}>Total Cost</th>
                      <th style={{ padding: "9px 12px", textAlign: "right", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Gross Profit</th>
                      <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Profit %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, idx) => {
                      const noMatch = r.boxCost === null;
                      return (
                        <tr key={idx} style={{ borderBottom: "1px solid #f1f5f9", background: noMatch ? "#fafafa" : idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
                          <td style={{ padding: "7px 12px", color: "#94a3b8", fontSize: 11 }}>{idx + 1}</td>
                          <td style={{ padding: "7px 12px", fontWeight: 700, color: noMatch ? "#94a3b8" : "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>
                            {r.code}
                            {noMatch && <span style={{ marginLeft: 5, fontSize: 9, color: "#f59e0b" }}>no BOM</span>}
                          </td>
                          <td style={{ padding: "7px 12px", textAlign: "right", direction: "rtl", color: "#374151", fontSize: 11, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.desc || <span style={{ color: "#94a3b8" }}>—</span>}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: r.qty < 0 ? "#dc2626" : "#374151" }}>{_fmtNum.format(r.qty)}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: r.amount < 0 ? "#dc2626" : "#374151" }}>{fmt(r.amount)}</td>
                          <td style={{ padding: "7px 12px", textAlign: "center", fontFamily: "monospace", color: "#64748b", fontSize: 11 }}>{r.conv ?? <span style={{ color: "#94a3b8" }}>—</span>}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: "#1d4ed8" }}>{noMatch ? <span style={{ color: "#94a3b8" }}>—</span> : fmt(r.boxCost)}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed", fontWeight: 600 }}>{noMatch ? <span style={{ color: "#94a3b8" }}>—</span> : fmtQty(r.unitCost)}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed", fontWeight: 700 }}>{noMatch ? <span style={{ color: "#94a3b8" }}>—</span> : fmt(r.totalCost)}</td>
                          <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: noMatch ? "#94a3b8" : r.grossProfit >= 0 ? "#15803d" : "#dc2626", fontWeight: 600 }}>
                            {noMatch ? "—" : fmt(r.grossProfit)}
                          </td>
                          <td style={{ padding: "7px 12px", textAlign: "center" }}>
                            {noMatch ? <span style={{ color: "#94a3b8" }}>—</span> : (
                              <span style={{ background: profitBg(r.profitPct), color: profitColor(r.profitPct), padding: "2px 8px", borderRadius: 20, fontWeight: 700, fontSize: 11, fontFamily: "monospace" }}>{fmtPct(r.profitPct)}%</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot style={{ position: "sticky", bottom: 0, background: "#f8fafc", borderTop: "2px solid #e2e8f0" }}>
                    <tr>
                      <td colSpan={5} style={{ padding: "9px 12px", fontWeight: 800, color: "#0f2d5a" }}>TOTAL</td>
                      <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#0f2d5a" }}>{fmt(totalAmount)}</td>
                      <td style={{ padding: "9px 12px" }}></td>
                      <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#1d4ed8" }}>{fmt(totalCostSum)}</td>
                      <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: totalProfit >= 0 ? "#15803d" : "#dc2626" }}>{fmt(totalProfit)}</td>
                      <td style={{ padding: "9px 12px", textAlign: "center" }}>
                        {totalAmount > 0 && <span style={{ background: profitBg(Math.round(totalProfit/totalAmount*100*10)/10), color: profitColor(Math.round(totalProfit/totalAmount*100*10)/10), padding: "2px 8px", borderRadius: 20, fontWeight: 800, fontSize: 11, fontFamily: "monospace" }}>{fmtPct(Math.round(totalProfit/totalAmount*100*10)/10)}%</span>}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          </div>
        );
      })()}
      {/* Finished Goods Page */}
      {page === "fg" && (() => {
        const costMap = Object.fromEntries(recalcData.filter(i => i.hasBOM).map(i => [i.code, { boxCost: i.boxCost, rawBoxCost: i.rawBoxCost, profitPct: i.profitPct, profit: i.profit }]));
        const rows = FINISHED_GOODS.map(r => {
          const cost = costMap[r.code] ?? null;
          return { ...r, ...( cost ? { ...cost, hasCost: true } : { hasCost: false }) };
        });
        const filtered = rows
          .filter(r => r.code.toLowerCase().includes(fgSearch.toLowerCase()) || (r.desc||'').includes(fgSearch))
          .filter(r => fgFilter === "all" ? true : fgFilter === "priced" ? r.sp !== null : fgFilter === "costed" ? r.hasCost : !r.hasCost);
        const withPrice = rows.filter(r => r.sp !== null).length;
        const withCost  = rows.filter(r => r.hasCost).length;
        return (
          <div style={{ padding: "16px 20px" }}>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 14 }}>
              {[
                { label: "Total Items",     value: rows.length,  icon: "📦", color: "#1d4ed8", bg: "#eff6ff" },
                { label: "With Sell Price", value: withPrice,    icon: "💰", color: "#15803d", bg: "#f0fdf4" },
                { label: "With BOM Cost",   value: withCost,     icon: "🏭", color: "#7c3aed", bg: "#f5f3ff" },
                { label: "No Cost Data",    value: rows.length - withCost, icon: "⚠️", color: "#b45309", bg: "#fffbeb" },
              ].map((s,i) => (
                <div key={i} style={{ background: s.bg, borderRadius: 10, padding: "12px 14px", border: `1px solid ${s.color}22` }}>
                  <div style={{ fontSize: 20 }}>{s.icon}</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>{s.label}</div>
                </div>
              ))}
            </div>

            <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
              {/* Toolbar */}
              <div style={{ padding: "10px 14px", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <input value={fgSearch} onChange={e => setFgSearch(e.target.value)} placeholder="🔍  Search by code or name..."
                  style={{ border: "1px solid #cbd5e1", borderRadius: 7, padding: "6px 12px", fontSize: 13, fontFamily: "inherit", flex: 1, minWidth: 200 }} />
                {[["all","All"],["priced","💰 With Price"],["costed","🏭 With Cost"],["nodata","❌ No Cost"]].map(([f,label]) => (
                  <button key={f} onClick={() => setFgFilter(f)}
                    style={{ padding: "6px 12px", borderRadius: 7, border: "1px solid "+(fgFilter===f?"#1d4ed8":"#cbd5e1"), background: fgFilter===f?"#1d4ed8":"#fff", color: fgFilter===f?"#fff":"#374151", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600 }}>
                    {label}
                  </button>
                ))}
                <button onClick={() => {
                  const wb = XLSX.utils.book_new();
                  const wsRows = [
                    ["#","Item Code","Item Description","Sell Price","Conv.","Material Cost","Total Cost +35%","Profit","Profit %"],
                    ...filtered.map((r,idx) => [idx+1, r.code, r.desc||"", r.sp??""  , r.conv??"", r.hasCost?r.rawBoxCost:"", r.hasCost?r.boxCost:"", r.hasCost?r.profit:"", r.hasCost?r.profitPct:""]),
                  ];
                  const ws = XLSX.utils.aoa_to_sheet(wsRows);
                  ws["!cols"] = [4,14,28,10,6,14,14,10,8].map(w=>({wch:w}));
                  XLSX.utils.book_append_sheet(wb, ws, "Finished Goods");
                  XLSX.writeFile(wb, "Finished_Goods.xlsx");
                }}
                  style={{ background: "#15803d", border: "none", color: "#fff", padding: "6px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit", whiteSpace: "nowrap" }}>
                  ⬇️ Export Excel
                </button>
                <span style={{ color: "#94a3b8", fontSize: 12 }}>{filtered.length} items</span>
              </div>

              <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 300px)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "9px 10px", textAlign: "left",   borderBottom: "2px solid #e2e8f0", color: "#475569" }}>#</th>
                      <th style={{ padding: "9px 10px", textAlign: "left",   borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Item Code</th>
                      <th style={{ padding: "9px 10px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Description</th>
                      <th style={{ padding: "9px 10px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#15803d" }}>Sell Price</th>
                      <th style={{ padding: "9px 10px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#64748b" }}>Conv.</th>
                      <th style={{ padding: "9px 10px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#7c3aed" }}>Material Cost</th>
                      <th style={{ padding: "9px 10px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#1d4ed8" }}>Total Cost +35%</th>
                      <th style={{ padding: "9px 10px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Profit</th>
                      <th style={{ padding: "9px 10px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Profit %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, idx) => (
                      <tr key={r.code} style={{ borderBottom: "1px solid #f1f5f9", background: !r.hasCost ? "#fafafa" : idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
                        <td style={{ padding: "7px 10px", color: "#94a3b8", fontSize: 11 }}>{idx + 1}</td>
                        <td style={{ padding: "7px 10px", fontWeight: 700, color: r.hasCost ? "#0f2d5a" : "#94a3b8", fontFamily: "monospace", fontSize: 11 }}>
                          {r.code}
                          {!r.hasCost && <span style={{ marginLeft: 5, fontSize: 9, color: "#f59e0b" }}>no BOM</span>}
                        </td>
                        <td style={{ padding: "7px 10px", textAlign: "right", direction: "rtl", color: "#374151", fontSize: 11 }}>{r.desc || "—"}</td>
                        <td style={{ padding: "7px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: r.sp ? "#15803d" : "#94a3b8" }}>{r.sp ? fmt(r.sp) : "—"}</td>
                        <td style={{ padding: "7px 10px", textAlign: "center", fontFamily: "monospace", color: "#64748b" }}>{r.conv ?? "—"}</td>
                        <td style={{ padding: "7px 10px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed" }}>{r.hasCost ? fmt(r.rawBoxCost) : <span style={{ color: "#94a3b8" }}>—</span>}</td>
                        <td style={{ padding: "7px 10px", textAlign: "right", fontFamily: "monospace", color: "#1d4ed8", fontWeight: 700 }}>{r.hasCost ? fmt(r.boxCost) : <span style={{ color: "#94a3b8" }}>—</span>}</td>
                        <td style={{ padding: "7px 10px", textAlign: "right", fontFamily: "monospace", color: r.hasCost ? (r.profit >= 0 ? "#15803d" : "#dc2626") : "#94a3b8", fontWeight: 600 }}>{r.hasCost ? fmt(r.profit) : "—"}</td>
                        <td style={{ padding: "7px 10px", textAlign: "center" }}>
                          {r.hasCost
                            ? <span style={{ background: profitBg(r.profitPct), color: profitColor(r.profitPct), padding: "2px 8px", borderRadius: 20, fontWeight: 700, fontSize: 11, fontFamily: "monospace" }}>{fmtPct(r.profitPct)}%</span>
                            : <span style={{ color: "#94a3b8" }}>—</span>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        );
      })()}
      {/* Item Balance Page */}
      {page === "balance" && (() => {
        const costMap    = Object.fromEntries(recalcData.filter(i => i.hasBOM).map(i => [i.code, { boxCost: i.boxCost, conv: i.conv }]));
        const rawSource  = apiRawCost.length > 0
          ? apiRawCost.map(r => ({ code: String(r.ItemCode||"").trim(), lastCost: r.LastCost||0 }))
          : RAW_MATERIALS.map(r => ({ code: r.code, lastCost: r.lastCost||0 }));
        const rawCostMap = Object.fromEntries(rawSource.map(r => [r.code, r.lastCost]));
        const warehouses = ["all", ...new Set(balanceData.map(r => r.Warehouse).filter(Boolean))].sort();

        // FG tab data
        const fgRows = balanceData
          .filter(r => r.ItemType === "F")
          .filter(r => balanceWarehouse === "all" || r.Warehouse === balanceWarehouse)
          .filter(r => !balanceSearch || (r.ItemCode||"").toLowerCase().includes(balanceSearch.toLowerCase()) || (r.ItemDescription||"").includes(balanceSearch))
          .map(r => {
            const code = String(r.ItemCode||"").trim();
            const bal  = parseFloat(r.ItemBalance) || 0;
            const conv = parseFloat(r.SellingConversion) || 1;
            const balInBoxes = Math.round(bal / conv * 1000) / 1000;
            const c = costMap[code];
            const itemConv = c?.conv || conv;
            const unitCost = c ? Math.round(c.boxCost / itemConv * 10000) / 10000 : null;
            const amount   = unitCost !== null ? Math.round(unitCost * bal * 100) / 100 : null;
            return { ...r, bal, balInBoxes, unitCost, amount };
          });

        // RM tab data (R + P)
        const rmRows = balanceData
          .filter(r => r.ItemType === "R" || r.ItemType === "P")
          .filter(r => balanceWarehouse === "all" || r.Warehouse === balanceWarehouse)
          .filter(r => !balanceSearch || (r.ItemCode||"").toLowerCase().includes(balanceSearch.toLowerCase()) || (r.ItemDescription||"").includes(balanceSearch))
          .map(r => {
            const code = String(r.ItemCode||"").trim();
            const bal  = parseFloat(r.ItemBalance) || 0;
            const conv = parseFloat(r.SellingConversion) || 1;
            const balInBoxes = Math.round(bal / conv * 1000) / 1000;
            const unitCost = rawCostMap[code] ?? null;
            const amount   = unitCost !== null ? Math.round(unitCost * bal * 100) / 100 : null;
            return { ...r, bal, balInBoxes, unitCost, amount };
          });

        const activeRows = balanceTab === "fg" ? fgRows : rmRows;
        const totalBal    = activeRows.reduce((s, r) => s + r.bal, 0);
        const totalBoxes  = activeRows.reduce((s, r) => s + r.balInBoxes, 0);
        const totalAmount = activeRows.filter(r => r.amount !== null).reduce((s, r) => s + r.amount, 0);

        const renderTable = (rows, showCostLabel) => (
          <BalanceGrid rows={rows} showCostLabel={showCostLabel} balSort={balSort} setBalSort={setBalSort} colFilters={colFilters} setColFilters={setColFilters} fmt={fmt} fmtNum={_fmtNum} />
        );

        return (
          <div style={{ padding: "16px 20px" }}>
            {/* Action bar */}
            <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
              <button onClick={fetchBalance} disabled={balanceLoading}
                style={{ background: balanceLoading?"#94a3b8":"#1d4ed8", color:"#fff", border:"none", padding:"9px 20px", borderRadius:8, cursor:balanceLoading?"not-allowed":"pointer", fontFamily:"inherit", fontSize:13, fontWeight:700 }}>
                {balanceLoading?"⏳ Loading...":"🔄 Load Balance"}
              </button>
              {balanceMsg && <span style={{ fontSize:12, color:balanceMsg.startsWith("✅")?"#15803d":"#dc2626", fontWeight:600 }}>{balanceMsg}</span>}
            </div>

            {balanceData.length > 0 && <>
              {/* Tabs */}
              <div style={{ display:"flex", gap:4, marginBottom:12 }}>
                {[["fg","🏭 Finished Goods",fgRows.length],["rm","📦 Raw & Packaging (R+P)",rmRows.length]].map(([t,label,cnt])=>(
                  <button key={t} onClick={()=>{setBalanceTab(t);setBalanceSearch("");}}
                    style={{ padding:"9px 20px", borderRadius:"8px 8px 0 0", border:"none", cursor:"pointer", fontFamily:"inherit", fontSize:13, fontWeight:700,
                      background:balanceTab===t?"#fff":"#e2e8f0", color:balanceTab===t?"#0f2d5a":"#64748b",
                      borderBottom:balanceTab===t?"2px solid #1d4ed8":"2px solid transparent" }}>
                    {label} <span style={{background:"#e2e8f0",color:"#475569",borderRadius:20,padding:"1px 8px",fontSize:11,marginLeft:6}}>{cnt}</span>
                  </button>
                ))}
              </div>

              {/* Filter bar */}
              <div style={{ background:"#fff", borderRadius:"0 14px 14px 14px", boxShadow:"0 2px 16px rgba(0,0,0,0.08)", overflow:"hidden" }}>
                <div style={{ padding:"10px 14px", borderBottom:"1px solid #e2e8f0", display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                  <input value={balanceSearch} onChange={e=>setBalanceSearch(e.target.value)}
                    placeholder="🔍  Search by code or name..."
                    style={{ border:"1px solid #cbd5e1", borderRadius:7, padding:"6px 12px", fontSize:13, fontFamily:"inherit", flex:1, minWidth:200 }} />
                  <select value={balanceWarehouse} onChange={e=>setBalanceWarehouse(e.target.value)}
                    style={{ border:"1px solid #cbd5e1", borderRadius:7, padding:"6px 10px", fontSize:13, fontFamily:"inherit", background:"#fff" }}>
                    {warehouses.map(w=><option key={w} value={w}>{w==="all"?"All Warehouses":w}</option>)}
                  </select>
                  <button onClick={()=>{
                    const wb = XLSX.utils.book_new();
                    const wsRows = [
                      ["#","Item Code","Description","Type","Conv.","Warehouse","Balance (Unit)","Balance (Box)","Unit Cost","Amount"],
                      ...activeRows.map((r,i)=>[i+1,r.ItemCode,r.ItemDescription||"",r.ItemType||"",r.SellingConversion||"",r.Warehouse||"",r.bal,r.balInBoxes,r.unitCost||"",r.amount||""])
                    ];
                    const ws = XLSX.utils.aoa_to_sheet(wsRows);
                    ws["!cols"]=[4,14,28,6,6,16,12,10,12,14].map(w=>({wch:w}));
                    XLSX.utils.book_append_sheet(wb,ws,balanceTab==="fg"?"FG Balance":"RM Balance");
                    XLSX.writeFile(wb,`Item_Balance_${balanceTab==="fg"?"FG":"RM"}.xlsx`);
                  }} style={{ background:"#15803d", border:"none", color:"#fff", padding:"6px 14px", borderRadius:7, cursor:"pointer", fontSize:12, fontWeight:700, fontFamily:"inherit", whiteSpace:"nowrap" }}>
                    ⬇️ Export Excel
                  </button>
                </div>

                {balanceTab === "fg" && renderTable(fgRows, "Unit Cost (+35%)")}
                {balanceTab === "rm" && renderTable(rmRows, "Last Cost")}
              </div>
            </>}

            {balanceData.length === 0 && !balanceLoading && (
              <div style={{ textAlign:"center", padding:"80px 20px", color:"#94a3b8" }}>
                <div style={{ fontSize:48, marginBottom:16 }}>📦</div>
                <div style={{ fontSize:16, fontWeight:600, marginBottom:8 }}>No Data Loaded</div>
                <div style={{ fontSize:13 }}>Click <b>🔄 Load Balance</b> to fetch live inventory data</div>
              </div>
            )}
          </div>
        );
      })()}

      {/* Raw Material Cost Page */}
      {page === "rawcost" && (() => {
        const [rcSearch, setRcSearch] = [balanceSearch, setBalanceSearch];
        // Use full API data if available, else fall back to embedded RAW_MATERIALS
        const sourceData = apiRawCost.length > 0
          ? apiRawCost.map(r => ({ code: r.ItemCode, desc: r.ItemDescription||"", itemType: r.ItemType, lastCost: r.LastCost||0, avgCost: r.AverageCost||0 }))
          : RAW_MATERIALS;
        const filtered = sourceData
          .filter(r => !rcSearch || r.code.toLowerCase().includes(rcSearch.toLowerCase()) || (r.desc||"").includes(rcSearch));
        const totalLastCost = filtered.reduce((s, r) => s + (r.lastCost||0), 0);
        return (
          <div style={{ padding: "16px 20px" }}>
            <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 2px 16px rgba(0,0,0,0.08)", overflow: "hidden" }}>
              <div style={{ padding: "10px 14px", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 8, alignItems: "center" }}>
                <input value={rcSearch} onChange={e => setRcSearch(e.target.value)}
                  placeholder="🔍  Search by code or name..."
                  style={{ border: "1px solid #cbd5e1", borderRadius: 7, padding: "6px 12px", fontSize: 13, fontFamily: "inherit", flex: 1, minWidth: 200 }} />
                <button onClick={() => {
                  const wb = XLSX.utils.book_new();
                  const wsRows = [
                    ["#", "Item Code", "Description", "Type", "Last Cost", "Average Cost"],
                    ...filtered.map((r, i) => [i+1, r.code, r.desc||"", r.itemType||"", r.lastCost||0, r.avgCost||0])
                  ];
                  const ws = XLSX.utils.aoa_to_sheet(wsRows);
                  ws["!cols"] = [4,16,30,6,12,12].map(w=>({wch:w}));
                  XLSX.utils.book_append_sheet(wb, ws, "Raw Material Cost");
                  XLSX.writeFile(wb, "Raw_Material_Cost.xlsx");
                }} style={{ background: "#15803d", border: "none", color: "#fff", padding: "6px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>
                  ⬇️ Export Excel
                </button>
                <span style={{ fontSize: 12, color: "#94a3b8" }}>{filtered.length} items</span>
              </div>
              <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 240px)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "9px 12px", textAlign: "left",   borderBottom: "2px solid #e2e8f0", color: "#475569" }}>#</th>
                      <th style={{ padding: "9px 12px", textAlign: "left",   borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Item Code</th>
                      <th style={{ padding: "9px 12px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#475569" }}>Description</th>
                      <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "2px solid #e2e8f0", color: "#64748b" }}>Type</th>
                      <th style={{ padding: "9px 12px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#dc2626" }}>Last Cost</th>
                      <th style={{ padding: "9px 12px", textAlign: "right",  borderBottom: "2px solid #e2e8f0", color: "#7c3aed" }}>Average Cost</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, idx) => (
                      <tr key={r.code} style={{ borderBottom: "1px solid #f1f5f9", background: idx % 2 === 1 ? "#f8fafc" : "#fff" }}>
                        <td style={{ padding: "7px 12px", color: "#94a3b8", fontSize: 11 }}>{idx+1}</td>
                        <td style={{ padding: "7px 12px", fontWeight: 700, color: "#0f2d5a", fontFamily: "monospace", fontSize: 11 }}>{r.code}</td>
                        <td style={{ padding: "7px 12px", textAlign: "right", direction: "rtl", color: "#374151", fontSize: 11 }}>{r.desc||"—"}</td>
                        <td style={{ padding: "7px 12px", textAlign: "center" }}>
                          <span style={{ background: r.itemType==="R"?"#f0fdf4":"#fef9c3", color: r.itemType==="R"?"#15803d":"#92400e", padding: "2px 7px", borderRadius: 6, fontSize: 10, fontWeight: 700 }}>{r.itemType||"—"}</span>
                        </td>
                        <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#dc2626" }}>{fmt(r.lastCost)}</td>
                        <td style={{ padding: "7px 12px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed" }}>{fmt(r.avgCost)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot style={{ position: "sticky", bottom: 0, background: "#f8fafc", borderTop: "2px solid #e2e8f0" }}>
                    <tr>
                      <td colSpan={4} style={{ padding: "9px 12px", fontWeight: 800, color: "#0f2d5a" }}>TOTAL ({filtered.length} items)</td>
                      <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 800, color: "#dc2626" }}>{fmt(totalLastCost)}</td>
                      <td style={{ padding: "9px 12px" }}></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Cash Flow Page */}
      {page === "cashflow" && (() => {
        const r = cfData?.List0?.[0] || cfData?.List?.[0] || null;
        const n = (v) => parseFloat(v) || 0;
        const fmt2 = (v) => new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n(v));
        const fmt0 = (v) => new Intl.NumberFormat("en-US", { minimumFractionDigits: 0,  maximumFractionDigits: 0  }).format(n(v));
        const pct  = (v) => n(v).toFixed(2) + "%";
        return (
          <div style={{ padding: "16px 20px", maxWidth: 1100 }}>
            {/* Action bar */}
            <div style={{ display: "flex", gap: 10, marginBottom: 16, alignItems: "center" }}>
              <button onClick={fetchCashFlow} disabled={cfLoading}
                style={{ background: cfLoading ? "#94a3b8" : "#1d4ed8", color: "#fff", border: "none", padding: "9px 20px", borderRadius: 8, cursor: cfLoading ? "not-allowed" : "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700 }}>
                {cfLoading ? "⏳ Loading..." : "🔄 Load Cash Flow"}
              </button>
              {cfMsg && <span style={{ fontSize: 12, color: cfMsg.startsWith("✅") ? "#15803d" : "#dc2626", fontWeight: 600 }}>{cfMsg}</span>}
            </div>

            {!r && !cfLoading && (
              <div style={{ textAlign: "center", padding: "80px 20px", color: "#94a3b8" }}>
                <div style={{ fontSize: 48, marginBottom: 16 }}>💵</div>
                <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>No Data Loaded</div>
                <div style={{ fontSize: 13 }}>Click <b>🔄 Load Cash Flow</b> to fetch live data</div>
              </div>
            )}

            {r && <>
              {/* ═══ Hero Header ═══ */}
              <div style={{ background: "linear-gradient(135deg, #0f2d5a 0%, #1e40af 50%, #3b82f6 100%)", borderRadius: 16, padding: "24px 28px", marginBottom: 20, position: "relative", overflow: "hidden" }}>
                <div style={{ position: "absolute", top: -40, right: -40, width: 160, height: 160, borderRadius: 80, background: "rgba(255,255,255,0.06)" }} />
                <div style={{ position: "absolute", bottom: -30, right: 60, width: 100, height: 100, borderRadius: 50, background: "rgba(255,255,255,0.04)" }} />
                <div style={{ fontSize: 11, fontWeight: 700, color: "rgba(255,255,255,0.5)", letterSpacing: 2, marginBottom: 4 }}>FINANCIAL OVERVIEW</div>
                <div style={{ color: "#fff", fontSize: 24, fontWeight: 900 }}>Cash Flow — {r.Year} / {r.Month}</div>
                <div style={{ display: "flex", gap: 24, marginTop: 16 }}>
                  {[
                    { l: "Cash State", v: fmt2(r.CashState), c: n(r.CashState) >= 0 ? "#4ade80" : "#fca5a5" },
                    { l: "Available Cash", v: fmt2(n(r.TreasuryCashEGP) + n(r.BankCashEGP)), c: "#93c5fd" },
                    { l: "Sales Growth", v: n(r.YTDSalesGrowthPct).toFixed(2) + "%", c: n(r.YTDSalesGrowthPct) >= 0 ? "#4ade80" : "#fca5a5" },
                  ].map((m, i) => (
                    <div key={i}>
                      <div style={{ fontSize: 10, fontWeight: 600, color: "rgba(255,255,255,0.5)", letterSpacing: 0.5, marginBottom: 2 }}>{m.l}</div>
                      <div style={{ fontSize: 20, fontWeight: 900, fontFamily: "monospace", color: m.c }}>{m.v}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* ═══ Receivable VS Payable ═══ */}
              <CFVSGauge
                leftVal={n(r.TotalCollection)} rightVal={n(r.TotalCashPayable) + n(r.TotalTransferPayable)}
                leftLabel="TOTAL RECEIVABLE" rightLabel="TOTAL PAYABLE"
              />

              {/* ═══ Cash Position ═══ */}
              <CFSectionTitle icon="🏦" title="Cash Position" accent="#1d4ed8" />
              <CFGrid cols={3}>
                <CFMetricCard label="Opening Balance" value={fmt0(n(r.TreasuryCashEGP_Open) + n(r.BankCashEGP_Open))}
                  sub={`Treasury: ${fmt2(r.TreasuryCashEGP_Open)} | Bank: ${fmt2(r.BankCashEGP_Open)}`} color="#1d4ed8" icon="📂" />
                <CFMetricCard label="Cash State" value={fmt2(r.CashState)}
                  sub="Net cash flow position" color={n(r.CashState) >= 0 ? "#16a34a" : "#dc2626"} icon={n(r.CashState) >= 0 ? "✅" : "⚠️"} />
                <CFMetricCard label="Available Cash" value={fmt2(n(r.TreasuryCashEGP) + n(r.BankCashEGP))}
                  sub={`Treasury: ${fmt2(r.TreasuryCashEGP)} | Bank: ${fmt2(r.BankCashEGP)}`} color="#16a34a" icon="💰" />
              </CFGrid>

              {/* ═══ Sales & Collection ═══ */}
              <CFSectionTitle icon="📈" title="Sales & Collection" accent="#16a34a" />
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
                <CFGrid cols={2}>
                  <CFMetricCard label="Customer Invoices" value={fmt2(r.TotalCustomerSales)} sub="Total invoiced this period" color="#1d4ed8" icon="📄" />
                  <CFMetricCard label="Customer Payments" value={fmt2(r.TotalCustomerPayment)} sub="Collected this period" color="#16a34a" icon="💳" />
                </CFGrid>
                <CFMetricCard label="Collection Ratio"
                  value={pct(n(r.TotalCustomerPayment) / (n(r.TotalCustomerSales) || 1) * 100)}
                  sub="Payments ÷ Invoices" color="#d97706" icon="📊" />
              </div>

              {/* ═══ Purchasing ═══ */}
              <CFSectionTitle icon="🛒" title="Purchasing" accent="#7c3aed" />
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
                <CFGrid cols={2}>
                  <CFMetricCard label="Vendor Invoices" value={fmt2(r.TotalVendorsInvoices)} sub="Total vendor invoices" color="#7c3aed" icon="📋" />
                  <CFMetricCard label="Vendor Payments" value={fmt2(r.TotalVendorsPayment)} sub="Paid this period" color="#ea580c" icon="💸" />
                </CFGrid>
                <CFMetricCard label="Payment Ratio"
                  value={pct(n(r.VendorPaymentRatio) * 100)} sub="Payments ÷ Invoices" color="#d97706" icon="⚖️" />
              </div>

              {/* ═══ Expenses ═══ */}
              <CFSectionTitle icon="📉" title="Expenses" accent="#dc2626" />
              <CFGrid cols={3}>
                <CFMetricCard label="Total Expenses" value={fmt2(r.Expenses)} sub="Operating expenses" color="#dc2626" icon="🔻" />
                <CFMetricCard label="Expenses Ratio" value={pct(n(r.ExpensesRatio))} sub="Expenses ÷ Sales" color="#d97706" icon="📊" />
                <CFMetricCard label="Reference Sales" value={fmt2(r.TotalCustomerSales)} sub="For ratio calculation" color="#64748b" icon="📄" />
              </CFGrid>

              {/* ═══ Checks ═══ */}
              <CFSectionTitle icon="📝" title="Checks" accent="#0891b2" />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <CFMiniBar items={[
                  { label: "Check Collection", value: n(r.TotalCheckCollection), color: "#16a34a" },
                  { label: "Check Paid", value: n(r.TotalCheckPaid), color: "#dc2626" },
                  { label: "Due Checks", value: n(r.TotalDueCheck), color: "#d97706" },
                ]} />
                <CFMetricCard label="Net Check Position"
                  value={fmt2(n(r.TotalCheckCollection) - n(r.TotalCheckPaid))}
                  sub={`Due: ${fmt2(r.TotalDueCheck)}`}
                  color={n(r.TotalCheckCollection) - n(r.TotalCheckPaid) >= 0 ? "#16a34a" : "#dc2626"} icon="📝" />
              </div>

              {/* ═══ Sales Compare (YoY) ═══ */}
              <CFSectionTitle icon="📊" title="Sales Compare — Year over Year" accent="#d97706" />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div style={{ background: "#fff", borderRadius: 12, padding: "20px", border: "1px solid #e2e8f0" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 16 }}>
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 1 }}>YTD 2025</div>
                      <div style={{ fontSize: 22, fontWeight: 900, fontFamily: "monospace", color: "#64748b" }}>{fmt2(r.YTDSales2025)}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 1 }}>YTD 2026</div>
                      <div style={{ fontSize: 22, fontWeight: 900, fontFamily: "monospace", color: "#1d4ed8" }}>{fmt2(r.YTDSales2026)}</div>
                    </div>
                  </div>
                  {/* comparison bars */}
                  {(() => {
                    const max = Math.max(n(r.YTDSales2025), n(r.YTDSales2026), 1);
                    return <>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", width: 32 }}>2025</span>
                        <div style={{ flex: 1, height: 14, borderRadius: 7, background: "#f1f5f9", overflow: "hidden" }}>
                          <div style={{ height: "100%", borderRadius: 7, background: "#94a3b8", width: `${(n(r.YTDSales2025) / max * 100)}%` }} />
                        </div>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: 10, fontWeight: 700, color: "#1d4ed8", width: 32 }}>2026</span>
                        <div style={{ flex: 1, height: 14, borderRadius: 7, background: "#dbeafe", overflow: "hidden" }}>
                          <div style={{ height: "100%", borderRadius: 7, background: "linear-gradient(90deg, #1d4ed8, #3b82f6)", width: `${(n(r.YTDSales2026) / max * 100)}%` }} />
                        </div>
                      </div>
                    </>;
                  })()}
                </div>
                <CFMetricCard label="Sales Growth" value={n(r.YTDSalesGrowthPct).toFixed(2) + "%"}
                  sub="Year-over-year change" color={n(r.YTDSalesGrowthPct) >= 0 ? "#16a34a" : "#dc2626"} icon={n(r.YTDSalesGrowthPct) >= 0 ? "🚀" : "📉"}
                  trend={n(r.YTDSalesGrowthPct)} />
              </div>

              {/* ═══ Customer Balance ═══ */}
              <CFSectionTitle icon="👤" title="Customer Balance" accent="#d97706" />
              <CFGrid cols={3}>
                <CFMetricCard label="Open Balance" value={fmt2(r.CustomerOpenBalance)} color="#64748b" icon="📂" />
                <CFMetricCard label="Current Balance" value={fmt2(r.CustomerBalance)} color="#d97706" icon="💼" />
                <CFMetricCard label="Balance Growth" value={n(r.CustomerBalanceGrowth).toFixed(2) + "%"}
                  color={n(r.CustomerBalanceGrowth) >= 0 ? "#dc2626" : "#16a34a"} icon="📈" trend={n(r.CustomerBalanceGrowth)} />
              </CFGrid>

              {/* ═══ Vendor Balance ═══ */}
              <CFSectionTitle icon="🏭" title="Vendor Balance" accent="#7c3aed" />
              <CFGrid cols={3}>
                <CFMetricCard label="Open Balance" value={fmt2(r.VendorOpenBalance)} color="#64748b" icon="📂" />
                <CFMetricCard label="Current Balance" value={fmt2(r.VendorBalance)} color="#7c3aed" icon="🏢" />
                <CFMetricCard label="Balance Growth" value={n(r.VendorBalanceGrowth).toFixed(2) + "%"}
                  color={n(r.VendorBalanceGrowth) <= 0 ? "#16a34a" : "#dc2626"} icon="📊" trend={-n(r.VendorBalanceGrowth)} />
              </CFGrid>

              {/* ═══ Modern Sales ═══ */}
              <CFSectionTitle icon="🏪" title="Sales Channels" accent="#0891b2" />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <CFMiniBar items={[
                  { label: "Modern Trade", value: n(r.CustomerModernSales), color: "#0891b2" },
                  { label: "Wholesale", value: n(r.WholeSales), color: "#7c3aed" },
                  { label: "Yearly Collections", value: n(r.CustomerPaymentYearly), color: "#16a34a" },
                ]} />
                <div style={{ background: "#fff", borderRadius: 12, padding: "20px", border: "1px solid #e2e8f0" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 1, marginBottom: 12, textTransform: "uppercase" }}>Channel Split</div>
                  {(() => {
                    const mod = n(r.CustomerModernSales);
                    const ws = n(r.WholeSales);
                    const total = mod + ws || 1;
                    const modPct = (mod / total * 100).toFixed(1);
                    const wsPct = (ws / total * 100).toFixed(1);
                    return <>
                      <div style={{ display: "flex", height: 28, borderRadius: 14, overflow: "hidden", marginBottom: 12 }}>
                        <div style={{ width: `${modPct}%`, background: "linear-gradient(90deg, #0891b2, #22d3ee)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          <span style={{ fontSize: 10, fontWeight: 800, color: "#fff" }}>{modPct}%</span>
                        </div>
                        <div style={{ width: `${wsPct}%`, background: "linear-gradient(90deg, #7c3aed, #a78bfa)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          <span style={{ fontSize: 10, fontWeight: 800, color: "#fff" }}>{wsPct}%</span>
                        </div>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11 }}>
                        <span><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 4, background: "#0891b2", marginRight: 6 }} />Modern: {fmt2(mod)}</span>
                        <span><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 4, background: "#7c3aed", marginRight: 6 }} />Wholesale: {fmt2(ws)}</span>
                      </div>
                    </>;
                  })()}
                </div>
              </div>

              {/* ═══ Cash Breakdown ═══ */}
              <CFSectionTitle icon="💵" title="Cash Breakdown" accent="#16a34a" />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div style={{ background: "linear-gradient(135deg, #f0fdf4, #dcfce7)", borderRadius: 12, padding: "20px", border: "1px solid #bbf7d0" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#15803d", letterSpacing: 1, marginBottom: 14 }}>💰 INFLOW</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <div><div style={{ fontSize: 10, color: "#86efac", fontWeight: 600 }}>Cash</div><div style={{ fontSize: 18, fontWeight: 900, fontFamily: "monospace", color: "#15803d" }}>{fmt2(r.TotalCashReceivable)}</div></div>
                    <div><div style={{ fontSize: 10, color: "#86efac", fontWeight: 600 }}>Transfer</div><div style={{ fontSize: 18, fontWeight: 900, fontFamily: "monospace", color: "#15803d" }}>{fmt2(r.TotalTransferReceivable)}</div></div>
                  </div>
                </div>
                <div style={{ background: "linear-gradient(135deg, #fef2f2, #fee2e2)", borderRadius: 12, padding: "20px", border: "1px solid #fecaca" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#dc2626", letterSpacing: 1, marginBottom: 14 }}>🔻 OUTFLOW</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <div><div style={{ fontSize: 10, color: "#fca5a5", fontWeight: 600 }}>Cash</div><div style={{ fontSize: 18, fontWeight: 900, fontFamily: "monospace", color: "#dc2626" }}>{fmt2(r.TotalCashPayable)}</div></div>
                    <div><div style={{ fontSize: 10, color: "#fca5a5", fontWeight: 600 }}>Transfer</div><div style={{ fontSize: 18, fontWeight: 900, fontFamily: "monospace", color: "#dc2626" }}>{fmt2(r.TotalTransferPayable)}</div></div>
                  </div>
                </div>
              </div>
            </>}
          </div>
        );
      })()}

      {page === "pricelist" && (() => {
        try {
        const rows = plData?.List0 || [];
        const priceTypes = [...new Set(rows.map(r => r.PriceTypeDescription).filter(Boolean))].sort();
        const disc = parseFloat(plDiscount) || 25;

        // Require a price type to be selected — show prompt if not
        const selectedPT = plPriceType;

        // Filter rows for selected price type + search
        const fRows = rows.filter(r => {
          if (selectedPT && r.PriceTypeDescription !== selectedPT) return false;
          if (plSearch) {
            const s = plSearch.toLowerCase();
            return (r.ItemCode || "").toLowerCase().includes(s) || (r.ItemDescription || "").includes(plSearch);
          }
          return true;
        });

        // Build cost lookup
        const costMap = {};
        recalcData.forEach(r => { costMap[r.code] = { rawBoxCost: r.rawBoxCost, boxCost: r.boxCost }; });

        // Build flat item list for selected price type
        const items = fRows.map(r => {
          const sp = parseFloat(r.PriceSellingUnit) || 0;
          const costs = costMap[r.ItemCode] || {};
          const rawBoxCost = costs.rawBoxCost || 0;
          const boxCost = costs.boxCost || 0;
          const profit = sp && boxCost ? sp - boxCost : null;
          const profitPct = profit !== null && sp ? profit / sp * 100 : null;
          const netSP = sp * (1 - disc / 100);
          const netProfit = netSP && boxCost ? netSP - boxCost : null;
          const netProfitPct = netProfit !== null && netSP ? netProfit / netSP * 100 : null;
          return { ItemCode: r.ItemCode || "", ItemDescription: r.ItemDescription || "", conv: r.SellingConversion, sp, rawBoxCost, boxCost, profit, profitPct, netSP, netProfit, netProfitPct };
        });

        const pctColor = (v) => v === null ? "#cbd5e1" : v >= 40 ? "#15803d" : v >= 20 ? "#d97706" : "#dc2626";
        const pctBg    = (v) => v === null ? "transparent" : v >= 40 ? "#dcfce7" : v >= 20 ? "#fef3c7" : "#fef2f2";

        const TH = ({ children, align = "right", color = "#0f2d5a" }) => (
          <th style={{ padding: "10px 12px", textAlign: align, fontWeight: 700, color, borderBottom: "2px solid #e2e8f0", whiteSpace: "nowrap", fontSize: 11, background: "#f8fafc" }}>{children}</th>
        );

        return (
          <div>
            {/* Toolbar */}
            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}>
              <button onClick={fetchPriceList} disabled={plLoading}
                style={{ background: plLoading ? "#94a3b8" : "#1d4ed8", color: "#fff", border: "none", padding: "9px 18px", borderRadius: 8, cursor: plLoading ? "not-allowed" : "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 700 }}>
                {plLoading ? "⏳ Loading..." : "🔄 Load Price List"}
              </button>
              {plMsg && <span style={{ fontSize: 12, color: plMsg.startsWith("✅") ? "#15803d" : "#dc2626", fontWeight: 600 }}>{plMsg}</span>}

              {/* Price Type selector — prominent */}
              <select value={plPriceType} onChange={e => setPlPriceType(e.target.value)}
                style={{ padding: "8px 14px", border: "2px solid " + (plPriceType ? "#1d4ed8" : "#e2e8f0"), borderRadius: 8, fontSize: 13, fontFamily: "inherit", fontWeight: 700, color: plPriceType ? "#1d4ed8" : "#94a3b8", minWidth: 200 }}>
                <option value="">— Select Price Type —</option>
                {priceTypes.map(pt => <option key={pt} value={pt}>{pt}</option>)}
              </select>

              {/* Discount input */}
              <div style={{ display: "flex", alignItems: "center", gap: 6, background: "#fef3c7", border: "1px solid #fbbf24", borderRadius: 8, padding: "6px 12px" }}>
                <span style={{ fontSize: 12, fontWeight: 700, color: "#92400e" }}>Discount</span>
                <input type="number" min="0" max="100" step="0.5" value={plDiscount}
                  onChange={e => { const v = e.target.value; setPlDiscount(v === "" ? "" : parseFloat(v) || 0); }}
                  onBlur={e => { if (e.target.value === "" || isNaN(parseFloat(e.target.value))) setPlDiscount(25); }}
                  style={{ width: 52, padding: "2px 6px", border: "1px solid #fbbf24", borderRadius: 5, fontFamily: "monospace", fontSize: 13, fontWeight: 700, color: "#b45309", textAlign: "right" }} />
                <span style={{ fontSize: 12, fontWeight: 700, color: "#92400e" }}>%</span>
                {plDiscount !== 25 && <button onClick={() => setPlDiscount(25)} style={{ background: "none", border: "none", cursor: "pointer", color: "#94a3b8", fontSize: 11, padding: 0 }}>↺</button>}
              </div>

              <input placeholder="🔍 Search..." value={plSearch} onChange={e => setPlSearch(e.target.value)}
                style={{ padding: "8px 14px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, fontFamily: "inherit", width: 200 }} />

              {items.length > 0 && <span style={{ fontSize: 12, color: "#64748b", fontWeight: 600 }}>{items.length} items</span>}
            </div>

            {/* No data state */}
            {!plData && !plLoading && (
              <div style={{ textAlign: "center", padding: 60, color: "#94a3b8" }}>
                <div style={{ fontSize: 40, marginBottom: 10 }}>🏷️</div>
                <div style={{ fontWeight: 700 }}>Click "🔄 Load Price List" to fetch data</div>
              </div>
            )}

            {/* No price type selected */}
            {plData && !selectedPT && (
              <div style={{ textAlign: "center", padding: 60, color: "#64748b" }}>
                <div style={{ fontSize: 40, marginBottom: 10 }}>☝️</div>
                <div style={{ fontWeight: 700, fontSize: 15 }}>Select a Price Type to view the grid</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center", marginTop: 16 }}>
                  {priceTypes.map(pt => (
                    <button key={pt} onClick={() => setPlPriceType(pt)}
                      style={{ padding: "8px 16px", background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 600 }}>
                      {pt}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Grid */}
            {plData && selectedPT && (
              <>
              {/* ── Profitability KPIs ── */}
              {(() => {
                const withCost  = items.filter(i => i.boxCost > 0 && i.sp > 0);
                const total     = withCost.length;
                const avgPct    = total ? withCost.reduce((s, i) => s + (i.profitPct || 0), 0) / total : null;
                const above40   = withCost.filter(i => (i.profitPct || 0) >= 40);
                const range2040 = withCost.filter(i => (i.profitPct || 0) >= 20 && (i.profitPct || 0) < 40);
                const below20   = withCost.filter(i => (i.profitPct || 0) >= 0  && (i.profitPct || 0) < 20);
                const negative  = withCost.filter(i => (i.profitPct || 0) < 0);
                const totalSP   = withCost.reduce((s, i) => s + i.sp, 0);
                const totalCost = withCost.reduce((s, i) => s + i.boxCost, 0);
                const totalProfit = totalSP - totalCost;
                const best  = withCost.length ? withCost.reduce((a, b) => (a.profitPct||0) > (b.profitPct||0) ? a : b) : null;
                const worst = withCost.length ? withCost.reduce((a, b) => (a.profitPct||0) < (b.profitPct||0) ? a : b) : null;

                const KCard = ({ label, value, sub, color, icon, bg = "#fff" }) => (
                  <div style={{ background: bg, borderRadius: 12, padding: "16px 18px", border: "1px solid #e2e8f0", position: "relative", overflow: "hidden" }}>
                    <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 3, background: color }} />
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div>
                        <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: 1, textTransform: "uppercase", marginBottom: 6 }}>{label}</div>
                        <div style={{ fontSize: 22, fontWeight: 900, fontFamily: "monospace", color, lineHeight: 1.1 }}>{value}</div>
                        {sub && <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 5 }}>{sub}</div>}
                      </div>
                      <div style={{ fontSize: 22, opacity: 0.7 }}>{icon}</div>
                    </div>
                  </div>
                );

                // distribution bar buckets
                const buckets = [
                  { label: "≥ 40%", count: above40.length,   color: "#16a34a", bg: "#dcfce7" },
                  { label: "20–40%", count: range2040.length, color: "#d97706", bg: "#fef3c7" },
                  { label: "0–20%", count: below20.length,   color: "#ea580c", bg: "#ffedd5" },
                  { label: "< 0%",  count: negative.length,   color: "#dc2626", bg: "#fef2f2" },
                ];
                const maxBucket = Math.max(...buckets.map(b => b.count), 1);

                return (
                  <div style={{ marginBottom: 16 }}>
                    {/* Row 1: summary cards */}
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10, marginBottom: 10 }}>
                      <KCard label="Avg. Profit %" value={avgPct !== null ? avgPct.toFixed(1) + "%" : "—"} color={avgPct !== null && avgPct >= 40 ? "#15803d" : avgPct >= 20 ? "#d97706" : "#dc2626"} icon="📊" sub={`${total} items with cost`} />
                      <KCard label="Total Selling Value" value={fmt(totalSP)} color="#1d4ed8" icon="💰" sub={`${total} items`} />
                      <KCard label="Total Cost Value" value={fmt(totalCost)} color="#7c3aed" icon="🏭" sub="Sum of Cost +35%" />
                      <KCard label="Total Gross Profit" value={fmt(totalProfit)} color={totalProfit >= 0 ? "#15803d" : "#dc2626"} icon={totalProfit >= 0 ? "✅" : "⚠️"} sub={totalSP ? ((totalProfit / totalSP * 100).toFixed(1) + "% of sales") : ""} />
                    </div>

                    {/* Row 2: margin breakdown + best/worst */}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      {/* Distribution chart */}
                      <div style={{ background: "#fff", borderRadius: 12, padding: "16px 18px", border: "1px solid #e2e8f0" }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: "#0f2d5a", marginBottom: 12, letterSpacing: 0.5 }}>MARGIN DISTRIBUTION — {total} items</div>
                        {buckets.map(b => (
                          <div key={b.label} style={{ marginBottom: 10 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                              <span style={{ fontSize: 11, fontWeight: 700, color: b.color }}>{b.label}</span>
                              <span style={{ fontSize: 12, fontWeight: 800, fontFamily: "monospace", color: b.color }}>
                                {b.count} <span style={{ fontSize: 10, color: "#94a3b8" }}>({total ? (b.count / total * 100).toFixed(0) : 0}%)</span>
                              </span>
                            </div>
                            <div style={{ height: 8, borderRadius: 4, background: "#f1f5f9", overflow: "hidden" }}>
                              <div style={{ height: "100%", borderRadius: 4, background: b.color, width: `${b.count / maxBucket * 100}%`, transition: "width 0.4s ease" }} />
                            </div>
                          </div>
                        ))}
                      </div>

                      {/* Best / Worst */}
                      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                        <div style={{ background: "#f0fdf4", borderRadius: 12, padding: "14px 18px", border: "1px solid #bbf7d0", flex: 1 }}>
                          <div style={{ fontSize: 10, fontWeight: 700, color: "#15803d", letterSpacing: 1, marginBottom: 6 }}>🏆 BEST MARGIN</div>
                          {best ? <>
                            <div style={{ fontSize: 13, fontWeight: 800, fontFamily: "monospace", color: "#0f2d5a" }}>{best.ItemCode}</div>
                            <div style={{ fontSize: 11, color: "#64748b", marginTop: 2, marginBottom: 6 }}>{best.ItemDescription}</div>
                            <div style={{ display: "flex", gap: 10 }}>
                              <span style={{ fontSize: 18, fontWeight: 900, color: "#15803d", fontFamily: "monospace" }}>{(best.profitPct || 0).toFixed(1)}%</span>
                              <span style={{ fontSize: 11, color: "#94a3b8", alignSelf: "flex-end" }}>SP: {fmt(best.sp)}</span>
                            </div>
                          </> : <span style={{ color: "#94a3b8" }}>—</span>}
                        </div>
                        <div style={{ background: "#fef2f2", borderRadius: 12, padding: "14px 18px", border: "1px solid #fecaca", flex: 1 }}>
                          <div style={{ fontSize: 10, fontWeight: 700, color: "#dc2626", letterSpacing: 1, marginBottom: 6 }}>⚠️ WORST MARGIN</div>
                          {worst ? <>
                            <div style={{ fontSize: 13, fontWeight: 800, fontFamily: "monospace", color: "#0f2d5a" }}>{worst.ItemCode}</div>
                            <div style={{ fontSize: 11, color: "#64748b", marginTop: 2, marginBottom: 6 }}>{worst.ItemDescription}</div>
                            <div style={{ display: "flex", gap: 10 }}>
                              <span style={{ fontSize: 18, fontWeight: 900, color: "#dc2626", fontFamily: "monospace" }}>{(worst.profitPct || 0).toFixed(1)}%</span>
                              <span style={{ fontSize: 11, color: "#94a3b8", alignSelf: "flex-end" }}>SP: {fmt(worst.sp)}</span>
                            </div>
                          </> : <span style={{ color: "#94a3b8" }}>—</span>}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* ── Table ── */}
              <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid #e2e8f0", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "inherit" }}>
                  <thead>
                    <tr>
                      <TH align="left" color="#64748b">#</TH>
                      <TH align="left" color="#0f2d5a">Item Code</TH>
                      <TH align="left" color="#0f2d5a">Description</TH>
                      <TH align="center" color="#0f2d5a">Conv.</TH>
                      <TH color="#0f2d5a">Selling Price</TH>
                      <TH color="#7c3aed">Material Cost</TH>
                      <TH color="#7c3aed">Total Cost +35%</TH>
                      <TH color="#16a34a">Profit</TH>
                      <TH color="#16a34a">Profit %</TH>
                      <TH color="#d97706">Net Price ({disc}% off)</TH>
                      <TH color="#d97706">Profit % after Disc.</TH>
                    </tr>
                  </thead>
                  <tbody>
                    {items.length === 0 && (
                      <tr><td colSpan={11} style={{ padding: 30, textAlign: "center", color: "#94a3b8" }}>No items found</td></tr>
                    )}
                    {items.map((it, i) => {
                      const bg = i % 2 === 0 ? "#fff" : "#f8fafc";
                      return (
                        <tr key={it.ItemCode + i} style={{ background: bg, borderBottom: "1px solid #f1f5f9" }}>
                          <td style={{ padding: "8px 12px", color: "#94a3b8", fontSize: 11 }}>{i + 1}</td>
                          <td style={{ padding: "8px 12px", fontWeight: 700, fontFamily: "monospace", fontSize: 12, color: "#1d4ed8", whiteSpace: "nowrap" }}>{it.ItemCode}</td>
                          <td style={{ padding: "8px 12px", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{it.ItemDescription}</td>
                          <td style={{ padding: "8px 12px", textAlign: "center", fontFamily: "monospace" }}>{it.conv}</td>
                          {/* Selling Price */}
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#0f2d5a" }}>{it.sp ? fmt(it.sp) : "—"}</td>
                          {/* Material Cost */}
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", color: "#7c3aed" }}>{it.rawBoxCost ? fmt(it.rawBoxCost) : "—"}</td>
                          {/* Total Cost +35% */}
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#7c3aed" }}>{it.boxCost ? fmt(it.boxCost) : "—"}</td>
                          {/* Profit */}
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: it.profit === null ? "#cbd5e1" : it.profit >= 0 ? "#15803d" : "#dc2626" }}>
                            {it.profit === null ? "—" : fmt(it.profit)}
                          </td>
                          {/* Profit % */}
                          <td style={{ padding: "8px 12px", textAlign: "center" }}>
                            {it.profitPct === null ? <span style={{ color: "#cbd5e1" }}>—</span> :
                              <span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 12, fontSize: 11, fontWeight: 700, background: pctBg(it.profitPct), color: pctColor(it.profitPct) }}>
                                {it.profitPct.toFixed(1)}%
                              </span>
                            }
                          </td>
                          {/* Net Price */}
                          <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: "#d97706" }}>
                            {it.sp ? fmt(it.netSP) : "—"}
                          </td>
                          {/* Profit % after Discount */}
                          <td style={{ padding: "8px 12px", textAlign: "center" }}>
                            {it.netProfitPct === null ? <span style={{ color: "#cbd5e1" }}>—</span> :
                              <span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 12, fontSize: 11, fontWeight: 700, background: pctBg(it.netProfitPct), color: pctColor(it.netProfitPct) }}>
                                {it.netProfitPct.toFixed(1)}%
                              </span>
                            }
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              </>
            )}
          </div>
        );
        } catch(e) {
          return <div style={{padding:20,color:"#dc2626",fontWeight:700}}>❌ Render error: {String(e.message)}</div>;
        }
      })()}

      {/* ══════════════════════════════════════════
          SALES ANALYSIS PAGE
      ══════════════════════════════════════════ */}
      {page === "salesan" && (() => {
        try {
        const n  = v => parseFloat(v) || 0;
        const f2 = v => new Intl.NumberFormat("en-US", { minimumFractionDigits:2, maximumFractionDigits:2 }).format(n(v));
        const fq = v => new Intl.NumberFormat("en-US", { minimumFractionDigits:2, maximumFractionDigits:2 }).format(n(v));
        const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

        const allRows = saData || [];

        // saFilter: { item:["code",...], customer:["no",...], salesperson:["name",...], month:["1","3",...] }
        const hasFilter = dim => saFilter[dim] && saFilter[dim].length > 0;
        const inFilter  = (dim, val) => hasFilter(dim) && saFilter[dim].includes(String(val));

        const toggleFilter = (dim, val) => {
          setSaFilter(f => {
            const nf = { ...f };
            const arr = [...(nf[dim] || [])];
            const idx = arr.indexOf(String(val));
            if (idx >= 0) arr.splice(idx, 1); else arr.push(String(val));
            if (arr.length === 0) delete nf[dim]; else nf[dim] = arr;
            return nf;
          });
        };

        // Apply all active filters
        const filteredRows = allRows.filter(r => {
          if (hasFilter("item")        && !inFilter("item",        r.ItemCode))                                return false;
          if (hasFilter("customer")    && !inFilter("customer",    r.CustomerNo))                              return false;
          if (hasFilter("salesperson") && !inFilter("salesperson", r.SalesName||r.CustomerSalesPerson||"—"))  return false;
          if (hasFilter("month")       && !inFilter("month",       parseInt(r.InvoiceMonth)))                  return false;
          if (hasFilter("family")      && !inFilter("family",      r.FamilyDescription||"—"))                 return false;
          return true;
        });

        // Column visibility based on saView array
        const showItem     = saView.includes("item");
        const showCustomer = saView.includes("customer");
        const showSales    = saView.includes("salesperson");
        const showMonth    = saView.includes("month");
        const showFamily   = saView.includes("family");

        // Build flat rows: group by all selected dimensions
        const groupKey = r => {
          const parts = [];
          if (showItem)     parts.push("I:" + r.ItemCode);
          if (showCustomer) parts.push("C:" + r.CustomerNo);
          if (showSales)    parts.push("S:" + (r.SalesName||r.CustomerSalesPerson||"—"));
          if (showMonth)    parts.push("M:" + r.InvoiceMonth);
          if (showFamily)   parts.push("F:" + (r.FamilyDescription||"—"));
          return parts.join("|") || "ALL";
        };

        const grouped = {};
        filteredRows.forEach(r => {
          const k = groupKey(r);
          if (!grouped[k]) grouped[k] = {
            _key: k,
            ItemCode:          r.ItemCode,
            ItemDesc:          r.ItemDescription,
            CustomerNo:        r.CustomerNo,
            CustomerName:      r.CustomerName,
            SalesName:         r.SalesName || r.CustomerSalesPerson || "—",
            Month:             parseInt(r.InvoiceMonth),
            FamilyDescription: r.FamilyDescription || "—",
            amount: 0, qty: 0, qtyBox: 0
          };
          grouped[k].amount += n(r.Amount);
          grouped[k].qty    += n(r.Qty);
          grouped[k].qtyBox += n(r.QtyBox);
        });

        let tableData = Object.values(grouped);

        // Search
        if (saSearch) {
          const s = saSearch.toLowerCase();
          tableData = tableData.filter(r =>
            (r.ItemCode||"").toLowerCase().includes(s) ||
            (r.ItemDesc||"").includes(saSearch) ||
            (r.CustomerName||"").includes(saSearch) ||
            (r.SalesName||"").toLowerCase().includes(s) ||
            (r.FamilyDescription||"").toLowerCase().includes(s)
          );
        }

        // Sort
        tableData = [...tableData].sort((a,b) => {
          let v = 0;
          if (saSort.col === "amount")   v = n(b.amount) - n(a.amount);
          else if (saSort.col === "qty") v = n(b.qty) - n(a.qty);
          else if (saSort.col === "qtyBox")  v = n(b.qtyBox) - n(a.qtyBox);
          else if (saSort.col === "item")    v = (a.ItemCode||"").localeCompare(b.ItemCode||"");
          else if (saSort.col === "customer") v = (a.CustomerName||"").localeCompare(b.CustomerName||"");
          else if (saSort.col === "family")   v = (a.FamilyDescription||"").localeCompare(b.FamilyDescription||"");
          return saSort.dir === "asc" ? -v : v;
        });

        const totals = { amount: filteredRows.reduce((s,r)=>s+n(r.Amount),0), qty: filteredRows.reduce((s,r)=>s+n(r.Qty),0), qtyBox: filteredRows.reduce((s,r)=>s+n(r.QtyBox),0) };
        const maxBar = Math.max(...tableData.map(r=>r.amount),1);

        // Apply per-column filters (saColF)
        let displayData = saColF && Object.values(saColF).some(v=>v) ? tableData.filter(r =>
          Object.entries(saColF).every(([col,val]) => {
            if (!val) return true;
            const v = val.toLowerCase();
            if (col==="item")     return (r.ItemCode||"").toLowerCase().includes(v);
            if (col==="itemdesc") return (r.ItemDesc||"").toLowerCase().includes(v);
            if (col==="customer") return (r.CustomerName||"").toLowerCase().includes(v)||(r.CustomerNo||"").toString().includes(v);
            if (col==="sales")    return (r.SalesName||"").toLowerCase().includes(v);
            if (col==="month")    return (MONTHS[(r.Month||1)-1]||"").toLowerCase().includes(v);
            if (col==="family")   return (r.FamilyDescription||"").toLowerCase().includes(v);
            if (col==="amount")   return String(Math.round(r.amount)).includes(v);
            if (col==="qty")      return String(Math.round(r.qty)).includes(v);
            if (col==="qtyBox")   return String(Math.round(r.qtyBox)).includes(v);
            return true;
          })
        ) : tableData;

        // Excel-style sortable TH with filter input — defined at module level as SalesSTH

        return (
          <div onClick={() => { setSaOpenMonth(false); setSaOpenGroup(false); setSaOpenCustomer(false); }} style={{ position:"relative" }}>
            {/* ── Toolbar ── */}
            <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap", marginBottom:12 }}>
              {/* Year picker */}
              <div style={{ display:"flex", alignItems:"center", gap:2, background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, padding:"4px 8px" }}>
                <button onClick={()=>{setSaYear(y=>y-1);setSaFilter({});setSaSelected(null);}} style={{ background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b",padding:"0 4px" }}>‹</button>
                <span style={{ fontWeight:800,fontSize:15,color:"#0f2d5a",minWidth:42,textAlign:"center" }}>{saYear}</span>
                <button onClick={()=>{setSaYear(y=>y+1);setSaFilter({});setSaSelected(null);}} style={{ background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b",padding:"0 4px" }}>›</button>
              </div>
              <button onClick={()=>fetchSales(saYear)} disabled={saLoading}
                style={{ background:saLoading?"#94a3b8":"#1d4ed8",color:"#fff",border:"none",padding:"8px 16px",borderRadius:8,cursor:saLoading?"not-allowed":"pointer",fontFamily:"inherit",fontSize:13,fontWeight:700 }}>
                {saLoading?"⏳ Loading...":"🔄 Load Sales"}
              </button>
              {saMsg && <span style={{ fontSize:12,color:saMsg.startsWith("✅")?"#15803d":"#dc2626",fontWeight:600 }}>{saMsg}</span>}
            </div>

            {/* ── Filters row ── */}
            <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:12, alignItems:"center" }}>

              {/* Month */}
              <details style={{ position:"relative", zIndex:10 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background:(saFilter.month||[]).length>0?"#dbeafe":"#f8fafc", border:`1px solid ${(saFilter.month||[]).length>0?"#93c5fd":"#d1d5db"}`, borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color:(saFilter.month||[]).length>0?"#1d4ed8":"#374151", userSelect:"none", whiteSpace:"nowrap" }}>
                  {(saFilter.month||[]).length>0 && <span style={{ fontSize:10,background:"#1d4ed8",color:"#fff",borderRadius:2,padding:"1px 5px",fontWeight:700 }}>{(saFilter.month||[]).length}</span>}
                  Month ▾
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:200 }}>
                  <div style={{ display:"flex", borderBottom:"1px solid #e5e7eb" }}>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFilter(f=>({...f,month:MONTHS.map((_,i)=>String(i+1))})); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#1d4ed8",fontWeight:700,textAlign:"center" }} onMouseEnter={e=>e.currentTarget.style.background="#eff6ff"} onMouseLeave={e=>e.currentTarget.style.background=""}>Select All</div>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFilter(f=>{const nf={...f};delete nf.month;return nf;}); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#dc2626",fontWeight:700,textAlign:"center",borderLeft:"1px solid #e5e7eb" }} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>e.currentTarget.style.background=""}>Clear</div>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr" }}>
                    {MONTHS.map((mn,i) => {
                      const val=String(i+1), active=(saFilter.month||[]).includes(val), hasData=allRows.some(r=>parseInt(r.InvoiceMonth)===(i+1));
                      return <div key={val} onClick={e=>{ if(hasData){ toggleFilter("month",i+1); } }} style={{ padding:"5px 10px",cursor:hasData?"pointer":"default",display:"flex",alignItems:"center",gap:6,background:active?"#eff6ff":"#fff",opacity:hasData?1:0.35,fontSize:12 }} onMouseEnter={e=>{ if(hasData&&!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background=active?"#eff6ff":"#fff"; }}>
                        <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#1d4ed8":"#9ca3af"}`,background:active?"#1d4ed8":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                        <span style={{ color:active?"#1d4ed8":"#374151",fontWeight:active?700:400 }}>{mn}</span>
                      </div>;
                    })}
                  </div>
                </div>
              </details>

              {/* Item */}
              <details style={{ position:"relative", zIndex:10 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background:(saFilter.item||[]).length>0?"#dbeafe":"#f8fafc", border:`1px solid ${(saFilter.item||[]).length>0?"#93c5fd":"#d1d5db"}`, borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color:(saFilter.item||[]).length>0?"#1d4ed8":"#374151", userSelect:"none", whiteSpace:"nowrap" }}>
                  {(saFilter.item||[]).length>0 && <span style={{ fontSize:10,background:"#1d4ed8",color:"#fff",borderRadius:2,padding:"1px 5px",fontWeight:700 }}>{(saFilter.item||[]).length}</span>}
                  Item ▾
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:260, maxHeight:300, display:"flex", flexDirection:"column" }}>
                  <div style={{ display:"flex", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFilter(f=>({...f,item:[...new Set(allRows.map(r=>r.ItemCode))]})); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#1d4ed8",fontWeight:700,textAlign:"center" }} onMouseEnter={e=>e.currentTarget.style.background="#eff6ff"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Select All</div>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFilter(f=>{const nf={...f};delete nf.item;return nf;}); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#dc2626",fontWeight:700,textAlign:"center",borderLeft:"1px solid #e5e7eb" }} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Clear</div>
                  </div>
                  <div style={{ overflowY:"auto", flex:1 }}>
                    {[...new Map(allRows.map(r=>[r.ItemCode,{code:r.ItemCode,desc:(r.ItemDescription||"").trim()}])).values()].sort((a,b)=>a.code.localeCompare(b.code)).map(it => {
                      const active=(saFilter.item||[]).includes(it.code);
                      return <div key={it.code} onClick={()=>toggleFilter("item",it.code)} style={{ padding:"5px 10px",cursor:"pointer",display:"flex",alignItems:"center",gap:6,background:active?"#eff6ff":"#fff",fontSize:12 }} onMouseEnter={e=>{ if(!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background="#fff"; }}>
                        <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#1d4ed8":"#9ca3af"}`,background:active?"#1d4ed8":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                        <div><div style={{ color:active?"#1d4ed8":"#374151",fontWeight:active?700:400,fontSize:11,fontFamily:"monospace" }}>{it.code}</div><div style={{ fontSize:10,color:"#9ca3af" }}>{it.desc}</div></div>
                      </div>;
                    })}
                  </div>
                </div>
              </details>

              {/* Item */}
              <details style={{ position:"relative", zIndex:10 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background:(saFilter.item||[]).length>0?"#dbeafe":"#f8fafc", border:`1px solid ${(saFilter.item||[]).length>0?"#93c5fd":"#d1d5db"}`, borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color:(saFilter.item||[]).length>0?"#1d4ed8":"#374151", userSelect:"none", whiteSpace:"nowrap" }}>
                  {(saFilter.item||[]).length>0 && <span style={{ fontSize:10,background:"#1d4ed8",color:"#fff",borderRadius:2,padding:"1px 5px",fontWeight:700 }}>{(saFilter.item||[]).length}</span>}
                  Item ▾
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:280, maxHeight:320, display:"flex", flexDirection:"column" }}>
                  <div style={{ padding:"6px 8px", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <input autoFocus value={saItemSearch} onChange={e=>setSaItemSearch(e.target.value)} onClick={e=>e.stopPropagation()} placeholder="🔍 Search item..." style={{ width:"100%", padding:"4px 8px", border:"1px solid #d1d5db", borderRadius:4, fontSize:11, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }} />
                  </div>
                  <div style={{ display:"flex", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaItemSearch(""); const all=[...new Set(allRows.map(r=>r.ItemCode))]; setSaFilter(f=>({...f,item:all})); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#1d4ed8",fontWeight:700,textAlign:"center" }} onMouseEnter={e=>e.currentTarget.style.background="#eff6ff"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Select All</div>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaItemSearch(""); setSaFilter(f=>{const nf={...f};delete nf.item;return nf;}); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#dc2626",fontWeight:700,textAlign:"center",borderLeft:"1px solid #e5e7eb" }} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Clear</div>
                  </div>
                  <div style={{ overflowY:"auto", flex:1 }}>
                    {[...new Map(allRows.map(r=>[r.ItemCode,{code:r.ItemCode,desc:(r.ItemDescription||"").trim()}])).values()].sort((a,b)=>a.code.localeCompare(b.code)).filter(it=>{ const q=saItemSearch.toLowerCase(); return !q||it.code.toLowerCase().includes(q)||it.desc.toLowerCase().includes(q); }).map(it => {
                      const active=(saFilter.item||[]).includes(it.code);
                      return <div key={it.code} onClick={()=>toggleFilter("item",it.code)} style={{ padding:"5px 10px",cursor:"pointer",display:"flex",alignItems:"center",gap:6,background:active?"#eff6ff":"#fff",fontSize:12 }} onMouseEnter={e=>{ if(!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background="#fff"; }}>
                        <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#1d4ed8":"#9ca3af"}`,background:active?"#1d4ed8":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                        <div><div style={{ color:active?"#1d4ed8":"#374151",fontWeight:active?700:400,fontSize:11,fontFamily:"monospace" }}>{it.code}</div><div style={{ fontSize:10,color:"#9ca3af" }}>{it.desc}</div></div>
                      </div>;
                    })}
                  </div>
                </div>
              </details>

              {/* Customer */}
              <details style={{ position:"relative", zIndex:10 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background:(saFilter.customer||[]).length>0?"#dbeafe":"#f8fafc", border:`1px solid ${(saFilter.customer||[]).length>0?"#93c5fd":"#d1d5db"}`, borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color:(saFilter.customer||[]).length>0?"#1d4ed8":"#374151", userSelect:"none", whiteSpace:"nowrap" }}>
                  {(saFilter.customer||[]).length>0 && <span style={{ fontSize:10,background:"#1d4ed8",color:"#fff",borderRadius:2,padding:"1px 5px",fontWeight:700 }}>{(saFilter.customer||[]).length}</span>}
                  Customer ▾
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:260, maxHeight:320, display:"flex", flexDirection:"column" }}>
                  <div style={{ padding:"6px 8px", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <input autoFocus value={saCustSearch} onChange={e=>setSaCustSearch(e.target.value)} onClick={e=>e.stopPropagation()} placeholder="🔍 Search customer..." style={{ width:"100%", padding:"4px 8px", border:"1px solid #d1d5db", borderRadius:4, fontSize:11, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }} />
                  </div>
                  <div style={{ display:"flex", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaCustSearch(""); const all=[...new Set(allRows.map(r=>String(r.CustomerNo)))]; setSaFilter(f=>({...f,customer:all})); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#1d4ed8",fontWeight:700,textAlign:"center" }} onMouseEnter={e=>e.currentTarget.style.background="#eff6ff"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Select All</div>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaCustSearch(""); setSaFilter(f=>{const nf={...f};delete nf.customer;return nf;}); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#dc2626",fontWeight:700,textAlign:"center",borderLeft:"1px solid #e5e7eb" }} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Clear</div>
                  </div>
                  <div style={{ overflowY:"auto", flex:1 }}>
                    {[...new Map(allRows.map(r=>[String(r.CustomerNo),{no:String(r.CustomerNo),name:(r.CustomerName||"").trim()}])).values()].sort((a,b)=>a.name.localeCompare(b.name)).filter(c=>{ const q=saCustSearch.toLowerCase(); return !q||c.name.toLowerCase().includes(q)||c.no.toLowerCase().includes(q); }).map(c => {
                      const active=(saFilter.customer||[]).includes(c.no);
                      return <div key={c.no} onClick={()=>toggleFilter("customer",c.no)} style={{ padding:"5px 10px",cursor:"pointer",display:"flex",alignItems:"center",gap:6,background:active?"#eff6ff":"#fff",fontSize:12 }} onMouseEnter={e=>{ if(!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background="#fff"; }}>
                        <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#1d4ed8":"#9ca3af"}`,background:active?"#1d4ed8":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                        <div><div style={{ color:active?"#1d4ed8":"#374151",fontWeight:active?700:400,fontSize:12 }}>{c.name}</div><div style={{ fontSize:10,color:"#9ca3af" }}>{c.no}</div></div>
                      </div>;
                    })}
                  </div>
                </div>
              </details>

              {/* Family */}
              <details style={{ position:"relative", zIndex:12 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background: hasFilter("family")?"#fdf4ff":"#f3f4f6", border:`1px solid ${hasFilter("family")?"#d946ef":"#d1d5db"}`, borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color: hasFilter("family")?"#a21caf":"#374151", userSelect:"none", whiteSpace:"nowrap" }}>
                  🏷️ Family {hasFilter("family") ? `(${saFilter.family.length})` : "▾"}
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:240, maxHeight:320, display:"flex", flexDirection:"column" }}>
                  <div style={{ padding:"6px 8px", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <input autoFocus value={saFamilySearch} onChange={e=>setSaFamilySearch(e.target.value)} onClick={e=>e.stopPropagation()} placeholder="🔍 Search family..." style={{ width:"100%", padding:"4px 8px", border:"1px solid #d1d5db", borderRadius:4, fontSize:11, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }} />
                  </div>
                  <div style={{ display:"flex", borderBottom:"1px solid #e5e7eb", flexShrink:0 }}>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFamilySearch(""); const all=[...new Set(allRows.map(r=>r.FamilyDescription||"—"))]; setSaFilter(f=>({...f,family:all})); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#1d4ed8",fontWeight:700,textAlign:"center" }} onMouseEnter={e=>e.currentTarget.style.background="#eff6ff"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Select All</div>
                    <div onClick={e=>{ e.currentTarget.closest('details').removeAttribute('open'); setSaFamilySearch(""); setSaFilter(f=>{const nf={...f};delete nf.family;return nf;}); }} style={{ flex:1,padding:"5px 10px",fontSize:11,cursor:"pointer",color:"#dc2626",fontWeight:700,textAlign:"center",borderLeft:"1px solid #e5e7eb" }} onMouseEnter={e=>e.currentTarget.style.background="#fef2f2"} onMouseLeave={e=>{ e.currentTarget.style.background=""; }}>Clear</div>
                  </div>
                  <div style={{ overflowY:"auto", flex:1 }}>
                    {[...new Set(allRows.map(r=>r.FamilyDescription||"—"))].sort((a,b)=>a.localeCompare(b)).filter(f=>{ const q=saFamilySearch.toLowerCase(); return !q||f.toLowerCase().includes(q); }).map(fam => {
                      const active=(saFilter.family||[]).includes(fam);
                      return <div key={fam} onClick={()=>toggleFilter("family",fam)} style={{ padding:"6px 10px",cursor:"pointer",display:"flex",alignItems:"center",gap:6,background:active?"#fdf4ff":"#fff",fontSize:12 }} onMouseEnter={e=>{ if(!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background="#fff"; }}>
                        <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#a21caf":"#9ca3af"}`,background:active?"#a21caf":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                        <div style={{ color:active?"#a21caf":"#374151",fontWeight:active?700:400,fontSize:12 }}>{fam}</div>
                      </div>;
                    })}
                  </div>
                </div>
              </details>

              {/* Group By */}
              <details style={{ position:"relative", zIndex:10 }} data-sa-dd>
                <summary onClick={e=>{ e.preventDefault(); document.querySelectorAll("details[data-sa-dd]").forEach(d=>{ if(d!==e.currentTarget.parentElement) d.removeAttribute("open"); }); e.currentTarget.parentElement.toggleAttribute("open"); }}
                  style={{ listStyle:"none", display:"flex", alignItems:"center", gap:5, padding:"4px 10px", background:"#dbeafe", border:"1px solid #93c5fd", borderRadius:3, cursor:"pointer", fontSize:12, fontWeight:600, color:"#1d4ed8", userSelect:"none", whiteSpace:"nowrap" }}>
                  <span style={{ fontSize:10,color:"#6b7280",fontWeight:400 }}>Group By:</span>
                  {saView.map(v=>({item:"Item",customer:"Customer",salesperson:"Sales",month:"Month",family:"Family"}[v])).join(" + ")} ▾
                </summary>
                <div style={{ position:"absolute", top:"calc(100% + 2px)", left:0, background:"#fff", border:"1px solid #d1d5db", boxShadow:"0 4px 16px rgba(0,0,0,0.15)", minWidth:180 }}>
                  {[["item","Per Item"],["customer","Per Customer"],["salesperson","Per Salesperson"],["month","Per Month"],["family","Per Family"]].map(([v,lb]) => {
                    const active=saView.includes(v);
                    return <div key={v} onClick={()=>setSaView(prev=>{ const next=active?prev.filter(x=>x!==v):[...prev,v]; return next.length===0?["item"]:next; })} style={{ padding:"6px 12px",cursor:"pointer",display:"flex",alignItems:"center",gap:8,background:active?"#eff6ff":"#fff",fontSize:12 }} onMouseEnter={e=>{ if(!active)e.currentTarget.style.background="#f9fafb"; }} onMouseLeave={e=>{ if(!active)e.currentTarget.style.background="#fff"; }}>
                      <span style={{ width:13,height:13,borderRadius:2,border:`1px solid ${active?"#1d4ed8":"#9ca3af"}`,background:active?"#1d4ed8":"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{active&&<span style={{ color:"#fff",fontSize:9,fontWeight:900 }}>✓</span>}</span>
                      <span style={{ color:active?"#1d4ed8":"#374151",fontWeight:active?700:400 }}>{lb}</span>
                    </div>;
                  })}
                </div>
              </details>

              <input placeholder="🔍 Search..." value={saSearch} onChange={e=>setSaSearch(e.target.value)}
                style={{ padding:"4px 8px", border:"1px solid #d1d5db", borderRadius:3, fontSize:12, fontFamily:"inherit", width:140 }} />
              {Object.keys(saFilter).length>0 && (
                <button onClick={()=>setSaFilter({})} style={{ padding:"4px 10px", background:"#fef2f2", border:"1px solid #fecaca", color:"#dc2626", borderRadius:3, cursor:"pointer", fontSize:11, fontWeight:700 }}>✕ Clear</button>
              )}
              <span style={{ fontSize:11, color:"#64748b" }}>{filteredRows.length} rows · {tableData.length} groups</span>
            </div>

            {/* ── KPI Cards ── */}
            <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:10,marginBottom:12 }}>
              {[
                { label:"Amount",      value:totals.amount,  fmt:f2,  color:"#1d4ed8", icon:"💰" },
                { label:"Qty",         value:totals.qty,     fmt:fq,  color:"#7c3aed", icon:"📦" },
                { label:"Qty Box",     value:totals.qtyBox,  fmt:fq,  color:"#0891b2", icon:"🗳️" },
                { label:"Groups",      value:tableData.length, fmt:v=>v, color:"#64748b", icon:"#️⃣" },
              ].map(({label,value,fmt,color,icon})=>(
                <div key={label} style={{ background:"#fff",borderRadius:10,padding:"14px 16px",border:"1px solid #e2e8f0",position:"relative",overflow:"hidden" }}>
                  <div style={{ position:"absolute",top:0,left:0,right:0,height:3,background:color }} />
                  <div style={{ display:"flex",justifyContent:"space-between" }}>
                    <div>
                      <div style={{ fontSize:10,fontWeight:700,color:"#94a3b8",letterSpacing:1,textTransform:"uppercase",marginBottom:5 }}>{label}</div>
                      <div style={{ fontSize:20,fontWeight:900,fontFamily:"monospace",color }}>{fmt(value)}</div>
                    </div>
                    <div style={{ fontSize:22,opacity:0.5 }}>{icon}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* ── Table + Detail Panel ── */}
            <div style={{ display:"grid", gridTemplateColumns: saSelected ? "1fr 340px" : "1fr", gap:12, alignItems:"start" }}>
            <div style={{ overflowX:"auto",borderRadius:10,border:"1px solid #e2e8f0",boxShadow:"0 1px 4px rgba(0,0,0,0.06)" }}>
              <table style={{ width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"inherit" }}>
                <thead>
                  <tr>
                    <th style={{ padding:0, borderBottom:"2px solid #e2e8f0", background:"#f8fafc" }}>
                      <div style={{ padding:"6px 8px 2px", fontSize:11, color:"#64748b", fontWeight:700 }}>#</div>
                      <div style={{ padding:"0 4px 4px", height:22 }} />
                    </th>
                    {showItem && <SalesSTH col="item" align="left" width={90} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>📦 Item Code</SalesSTH>}
                    {showItem && <SalesSTH col="itemdesc" align="left" width={140} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>Description</SalesSTH>}
                    {showCustomer && <SalesSTH col="customer" align="left" width={140} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>👤 Customer</SalesSTH>}
                    {showFamily   && <SalesSTH col="family"   align="left" width={130} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>🏷️ Family</SalesSTH>}
                    {showSales && <SalesSTH col="sales" align="left" width={100} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>🧑‍💼 Salesperson</SalesSTH>}
                    {showMonth && <SalesSTH col="month" align="left" width={70} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>📅 Month</SalesSTH>}
                    <SalesSTH col="amount" width={110} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>Amount</SalesSTH>
                    <SalesSTH col="qty" width={80} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>Qty</SalesSTH>
                    <SalesSTH col="qtyBox" width={80} saSort={saSort} setSaSort={setSaSort} saColF={saColF} setSaColF={setSaColF}>Qty Box</SalesSTH>
                    <th style={{ padding:0, borderBottom:"2px solid #e2e8f0", background:"#f8fafc", minWidth:60 }}>
                      <div style={{ padding:"6px 8px 2px", fontSize:11, color:"#0f2d5a", fontWeight:700, textAlign:"right" }}>Share%</div>
                      <div style={{ padding:"0 4px 4px", height:22 }} />
                    </th>
                    <th style={{ padding:0, borderBottom:"2px solid #e2e8f0", background:"#f8fafc", minWidth:90 }}>
                      <div style={{ padding:"6px 8px 2px", fontSize:11, color:"#0f2d5a", fontWeight:700 }}>Bar</div>
                      <div style={{ padding:"0 4px 4px", height:22 }} />
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {displayData.length===0 && (
                    <tr><td colSpan={20} style={{ padding:30,textAlign:"center",color:"#94a3b8" }}>No data matching current filters</td></tr>
                  )}
                  {displayData.map((r,i)=>{
                    const share = totals.amount ? (r.amount/totals.amount*100) : 0;
                    const barW  = (r.amount/maxBar*100).toFixed(1);
                    const isSelected = saSelected && saSelected.key === r._key;
                    const bg    = isSelected ? "#eff6ff" : i%2===0?"#fff":"#f8fafc";
                    const clickKey = r.ItemCode || r._key;
                    return (
                      <tr key={r._key} onClick={()=>setSaSelected(prev => prev && prev.key===r._key ? null : { key:r._key, itemCode:r.ItemCode, itemDesc:r.ItemDesc, customerNo:r.CustomerNo, customerName:r.CustomerName })}
                        style={{ background:bg, borderBottom:"1px solid #f1f5f9", cursor:"pointer", outline: isSelected?"2px solid #3b82f6":"none", outlineOffset:"-1px" }}>
                        <td style={{ padding:"7px 10px",color:"#94a3b8",fontSize:11 }}>{i+1}</td>
                        {showItem && <td style={{ padding:"7px 10px",fontWeight:700,fontFamily:"monospace",fontSize:11,color:"#1d4ed8",whiteSpace:"nowrap" }}>{r.ItemCode}</td>}
                        {showItem && <td style={{ padding:"7px 10px",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#374151" }}>{r.ItemDesc}</td>}
                        {showCustomer && <td style={{ padding:"7px 10px",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}><div style={{ fontWeight:600,fontSize:11 }}>{r.CustomerName}</div><div style={{ fontSize:10,color:"#94a3b8" }}>{r.CustomerNo}</div></td>}
                        {showFamily   && <td style={{ padding:"7px 10px",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#7c2d8e",fontWeight:500 }}>{r.FamilyDescription}</td>}
                        {showSales && <td style={{ padding:"7px 10px",whiteSpace:"nowrap",color:"#374151" }}>{r.SalesName}</td>}
                        {showMonth && <td style={{ padding:"7px 10px",fontWeight:700,color:"#d97706" }}>{MONTHS[(r.Month||1)-1]}</td>}
                        <td style={{ padding:"7px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#0f2d5a" }}>{f2(r.amount)}</td>
                        <td style={{ padding:"7px 10px",textAlign:"right",fontFamily:"monospace",color:"#7c3aed" }}>{fq(r.qty)}</td>
                        <td style={{ padding:"7px 10px",textAlign:"right",fontFamily:"monospace",color:"#0891b2" }}>{fq(r.qtyBox)}</td>
                        <td style={{ padding:"7px 10px",textAlign:"right",fontSize:11,color:"#64748b" }}>{share.toFixed(1)}%</td>
                        <td style={{ padding:"7px 10px" }}>
                          <div style={{ height:6,borderRadius:3,background:"#f1f5f9",overflow:"hidden" }}>
                            <div style={{ height:"100%",borderRadius:3,background:"#1d4ed8",width:barW+"%" }} />
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                  <tr style={{ background:"#f1f5f9",borderTop:"2px solid #e2e8f0" }}>
                    <td colSpan={2 + (showItem?2:0) + (showCustomer?1:0) + (showFamily?1:0) + (showSales?1:0) + (showMonth?1:0)}
                      style={{ padding:"9px 10px",fontWeight:800,color:"#0f2d5a",fontSize:12 }}>TOTAL ({displayData.length})</td>
                    <td style={{ padding:"9px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:"#1d4ed8" }}>{f2(totals.amount)}</td>
                    <td style={{ padding:"9px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:"#7c3aed" }}>{fq(totals.qty)}</td>
                    <td style={{ padding:"9px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:"#0891b2" }}>{fq(totals.qtyBox)}</td>
                    <td style={{ padding:"9px 10px",textAlign:"right",fontSize:11,color:"#64748b" }}>100%</td>
                    <td />
                  </tr>
                </tbody>
              </table>
            </div>

            {/* ── Detail Panel ── */}
            {saSelected && (() => {
              const selKey = saSelected.key;
              // Determine what was clicked: if group is by item, show customers for that item; else show items for that customer
              const isItemRow   = !!saSelected.itemCode && !saSelected.customerNo;
              const isCustomerRow = !!saSelected.customerNo && !saSelected.itemCode;
              const isMixedRow  = !!saSelected.itemCode && !!saSelected.customerNo;

              // Always show: customers who bought this item (from allRows, unfiltered by current filter)
              const panelItemCode = saSelected.itemCode;
              const panelCustNo   = saSelected.customerNo;

              // Rows from full dataset for this item
              const itemRows = panelItemCode
                ? allRows.filter(r => r.ItemCode === panelItemCode)
                : allRows.filter(r => String(r.CustomerNo) === String(panelCustNo));

              // Group by customer (if item selected) or by item (if customer selected)
              const groupMap = {};
              itemRows.forEach(r => {
                const gk = panelItemCode ? String(r.CustomerNo) : r.ItemCode;
                const gl = panelItemCode ? (r.CustomerName||"—") : (r.ItemDescription||r.ItemCode);
                if (!groupMap[gk]) groupMap[gk] = { key:gk, label:gl, amount:0, qty:0 };
                groupMap[gk].amount += parseFloat(r.Amount||r.amount||0);
                groupMap[gk].qty    += parseFloat(r.Qty||r.qty||0);
              });
              const groups = Object.values(groupMap).sort((a,b)=>b.amount-a.amount);
              const totalAmt = groups.reduce((s,g)=>s+g.amount,0);
              const panelTitle = panelItemCode
                ? (saSelected.itemDesc || panelItemCode)
                : (saSelected.customerName || panelCustNo);
              const panelSubtitle = panelItemCode ? panelItemCode : String(panelCustNo);
              const panelLabel = panelItemCode ? "Customers" : "Items";

              return (
                <div style={{ background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.1)", display:"flex", flexDirection:"column", maxHeight:"calc(100vh - 240px)", overflow:"hidden", position:"sticky", top:12 }}>
                  {/* Header */}
                  <div style={{ padding:"14px 16px", background:"linear-gradient(135deg,#0f2d5a,#1d4ed8)", color:"#fff" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"start" }}>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ fontWeight:800, fontSize:13, fontFamily:"monospace", marginBottom:2 }}>{panelSubtitle}</div>
                        <div style={{ fontSize:11, opacity:0.85, direction:"rtl", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{panelTitle}</div>
                      </div>
                      <button onClick={()=>setSaSelected(null)} style={{ background:"rgba(255,255,255,0.2)", border:"none", color:"#fff", width:26, height:26, borderRadius:6, cursor:"pointer", fontSize:14, flexShrink:0, marginLeft:8 }}>✕</button>
                    </div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginTop:10 }}>
                      <div style={{ background:"rgba(255,255,255,0.15)", borderRadius:7, padding:"7px 8px", textAlign:"center" }}>
                        <div style={{ fontSize:10, opacity:0.8 }}>Total Amount</div>
                        <div style={{ fontWeight:700, fontSize:13, fontFamily:"monospace" }}>{f2(totalAmt)}</div>
                      </div>
                      <div style={{ background:"rgba(255,255,255,0.15)", borderRadius:7, padding:"7px 8px", textAlign:"center" }}>
                        <div style={{ fontSize:10, opacity:0.8 }}>{panelLabel}</div>
                        <div style={{ fontWeight:700, fontSize:13, fontFamily:"monospace" }}>{groups.length}</div>
                      </div>
                    </div>
                  </div>
                  {/* List */}
                  <div style={{ overflowY:"auto", flex:1 }}>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                      <thead style={{ position:"sticky", top:0, zIndex:1 }}>
                        <tr style={{ background:"#f8fafc" }}>
                          <th style={{ padding:"6px 10px", textAlign:"left", borderBottom:"1px solid #e2e8f0", fontSize:11, color:"#64748b", fontWeight:700 }}>{panelLabel}</th>
                          <th style={{ padding:"6px 8px", textAlign:"right", borderBottom:"1px solid #e2e8f0", fontSize:11, color:"#64748b", fontWeight:700 }}>Amount</th>
                          <th style={{ padding:"6px 8px", textAlign:"right", borderBottom:"1px solid #e2e8f0", fontSize:11, color:"#64748b", fontWeight:700 }}>Share%</th>
                        </tr>
                      </thead>
                      <tbody>
                        {groups.map((g,idx)=>{
                          const share = totalAmt ? (g.amount/totalAmt*100) : 0;
                          const barW = (g.amount/(groups[0]?.amount||1)*100).toFixed(1);
                          return (
                            <tr key={g.key} style={{ borderBottom:"1px solid #f1f5f9", background: idx%2===0?"#fff":"#f8fafc" }}>
                              <td style={{ padding:"7px 10px" }}>
                                <div style={{ fontWeight:600, fontSize:11, color:"#0f2d5a", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:160 }}>{g.label}</div>
                                <div style={{ fontSize:9, color:"#94a3b8", fontFamily:"monospace" }}>{g.key}</div>
                                <div style={{ height:3, borderRadius:2, background:"#e2e8f0", marginTop:3, overflow:"hidden" }}>
                                  <div style={{ height:"100%", borderRadius:2, background:"#3b82f6", width:barW+"%" }} />
                                </div>
                              </td>
                              <td style={{ padding:"7px 8px", textAlign:"right", fontFamily:"monospace", fontWeight:700, color:"#1d4ed8", fontSize:11 }}>{f2(g.amount)}</td>
                              <td style={{ padding:"7px 8px", textAlign:"right", fontSize:11, color:"#64748b" }}>{share.toFixed(1)}%</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })()}
            </div>{/* end grid */}
          </div>
        );
        } catch(e) {
          return <div style={{padding:20,color:"#dc2626",fontWeight:700,fontSize:14}}>❌ Sales Error: {String(e.message)}</div>;
        }
      })()}

      {/* ── Price Lab ── */}
      {page === "pricelab" && (() => {
        try {
        // Reuse plData (Price List Type) as source — load if not loaded
        const rows = plData?.List0 || [];
        const priceTypes = [...new Set(rows.map(r => r.PriceTypeDescription).filter(Boolean))].sort();

        // Build cost map from BOM recalcData
        const costMap = {};
        recalcData.forEach(r => { costMap[r.code] = { rawBoxCost: r.rawBoxCost, boxCost: r.boxCost, hasBOM: r.hasBOM }; });

        // Filter by selected price type
        const ptRows = rows.filter(r => plabPriceType ? r.PriceTypeDescription === plabPriceType : false);

        // Build item list
        const items = ptRows.map(r => {
          const sp       = parseFloat(r.PriceSellingUnit) || 0;
          const costs    = costMap[r.ItemCode] || {};
          const hasCost  = !!(costs.boxCost);
          const boxCost  = costs.boxCost || 0;
          const rawBoxCost = costs.rawBoxCost || 0;
          const newSP    = parseFloat(plabPrices[r.ItemCode]) > 0 ? parseFloat(plabPrices[r.ItemCode]) : null;
          const disc     = parseFloat(plabDiscount) || 0;
          const curProfit    = sp && boxCost ? sp - boxCost : null;
          const curProfitPct = curProfit !== null && sp ? curProfit / sp * 100 : null;
          const newProfit    = newSP && boxCost ? newSP - boxCost : null;
          const newProfitPct = newProfit !== null && newSP ? newProfit / newSP * 100 : null;
          const netPrice     = disc > 0 && newSP ? Math.round(newSP * (1 - disc / 100) * 100) / 100 : null;
          const netProfit    = netPrice && boxCost ? netPrice - boxCost : null;
          const netProfitPct = netProfit !== null && netPrice ? netProfit / netPrice * 100 : null;
          return { code: r.ItemCode, desc: r.ItemDescription || "", conv: r.SellingConversion, sp, rawBoxCost, boxCost, hasCost, newSP, curProfit, curProfitPct, newProfit, newProfitPct, netPrice, netProfit, netProfitPct };
        }).filter(r => {
          if (!plabSearch) return true;
          const s = plabSearch.toLowerCase();
          return r.code.toLowerCase().includes(s) || r.desc.toLowerCase().includes(s);
        });

        // Traffic light
        const light = (pct) => pct === null ? null : pct >= 40 ? "🟢" : pct >= 20 ? "🟡" : "🔴";
        const pctColor = (v) => v === null ? "#94a3b8" : v >= 40 ? "#15803d" : v >= 20 ? "#d97706" : "#dc2626";
        const pctBg    = (v) => v === null ? "transparent" : v >= 40 ? "#dcfce7" : v >= 20 ? "#fef3c7" : "#fef2f2";

        // KPIs
        const withNewPrice = items.filter(r => r.newSP !== null && r.newSP > 0);
        const green  = withNewPrice.filter(r => r.newProfitPct >= 40).length;
        const yellow = withNewPrice.filter(r => r.newProfitPct >= 20 && r.newProfitPct < 40).length;
        const red    = withNewPrice.filter(r => r.newProfitPct < 20).length;
        const avgNew = withNewPrice.length ? withNewPrice.reduce((s,r) => s + (r.newProfitPct||0), 0) / withNewPrice.length : null;

        // Export Excel
        const exportLab = () => {
          const exp = withNewPrice.map(r => ({
            "Item Code":       r.code,
            "Description":     r.desc,
            "Conv.":           r.conv,
            "Current SP":      r.sp,
            "New Price":       r.newSP,
            "Material Cost":   r.rawBoxCost,
            "Total Cost +35%": r.boxCost,
            "New Profit":      r.newProfit,
            "New Profit %":    r.newProfitPct !== null ? +r.newProfitPct.toFixed(1) : null,
            "Status":          r.newProfitPct >= 40 ? "✅ Good" : r.newProfitPct >= 20 ? "⚠️ Marginal" : "❌ Low",
            ...(plabDiscount > 0 ? {
              [`Net Price (${plabDiscount}% off)`]: r.netPrice,
              "Net Profit":    r.netProfit,
              "Net Profit %":  r.netProfitPct !== null ? +r.netProfitPct.toFixed(1) : null,
            } : {}),
          }));
          const ws = XLSX.utils.json_to_sheet(exp);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Price Lab");
          XLSX.writeFile(wb, `PriceLab_${plabPriceType||"export"}.xlsx`);
        };

        return (
          <div style={{ padding:"0 20px 20px" }}>

            {/* Toolbar */}
            <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap", marginBottom:14 }}>
              <button onClick={fetchPriceList} disabled={plLoading}
                style={{ background: plLoading?"#94a3b8":"#1d4ed8", color:"#fff", border:"none", padding:"9px 18px", borderRadius:8, cursor: plLoading?"not-allowed":"pointer", fontFamily:"inherit", fontSize:13, fontWeight:700 }}>
                {plLoading ? "⏳ Loading..." : "🔄 Load Data"}
              </button>

              {/* Price Type selector */}
              {rows.length > 0 && (
                <select value={plabPriceType} onChange={e=>setPlabPriceType(e.target.value)}
                  style={{ padding:"8px 12px", border:"1px solid #d1d5db", borderRadius:8, fontSize:13, fontFamily:"inherit", background:"#fff", cursor:"pointer", minWidth:180 }}>
                  <option value="">— Select Price Type —</option>
                  {priceTypes.map(pt => <option key={pt} value={pt}>{pt}</option>)}
                </select>
              )}

              {plabPriceType && (
                <input placeholder="🔍 Search item..." value={plabSearch} onChange={e=>setPlabSearch(e.target.value)}
                  style={{ padding:"8px 12px", border:"1px solid #d1d5db", borderRadius:8, fontSize:13, fontFamily:"inherit", width:180 }} />
              )}

              {withNewPrice.length > 0 && (
                <button onClick={exportLab}
                  style={{ background:"#15803d", color:"#fff", border:"none", padding:"9px 16px", borderRadius:8, cursor:"pointer", fontFamily:"inherit", fontSize:13, fontWeight:700, marginLeft:"auto" }}>
                  📥 Export Excel ({withNewPrice.length})
                </button>
              )}

              {withNewPrice.length > 0 && (
                <button onClick={()=>setPlabPrices({})}
                  style={{ background:"#fef2f2", color:"#dc2626", border:"1px solid #fecaca", padding:"8px 14px", borderRadius:8, cursor:"pointer", fontFamily:"inherit", fontSize:13, fontWeight:700 }}>
                  ✕ Reset
                </button>
              )}

              {plMsg && <span style={{ fontSize:12, color: plMsg.startsWith("✅")?"#15803d":"#dc2626", fontWeight:600 }}>{plMsg}</span>}
            </div>

            {/* Discount row — always visible when price type selected */}
            {plabPriceType && (
              <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
                <div style={{ display:"flex", alignItems:"center", gap:6, background: plabDiscount > 0 ? "#fef3c7" : "#f8fafc", border:`1px solid ${plabDiscount > 0 ? "#fbbf24" : "#cbd5e1"}`, borderRadius:8, padding:"7px 14px" }}>
                  <span style={{ fontSize:13, color:"#475569", whiteSpace:"nowrap", fontWeight:600 }}>🏷️ Discount</span>
                  <input type="number" min="0" max="100" step="0.5" value={plabDiscount}
                    onChange={e => setPlabDiscount(Math.min(100, Math.max(0, parseFloat(e.target.value) || 0)))}
                    style={{ width:60, padding:"3px 8px", border:"1px solid #cbd5e1", borderRadius:5, fontFamily:"monospace", fontSize:14, fontWeight:700, color: plabDiscount > 0 ? "#b45309" : "#374151", textAlign:"right" }} />
                  <span style={{ fontSize:13, color:"#64748b" }}>%</span>
                  {plabDiscount > 0 && <button onClick={()=>setPlabDiscount(0)} style={{ background:"none", border:"none", cursor:"pointer", color:"#94a3b8", fontSize:13, padding:0 }}>✕</button>}
                </div>
                {plabDiscount > 0 && <span style={{ fontSize:12, color:"#b45309", fontWeight:600 }}>Net Price and Net Profit% columns will appear →</span>}
              </div>
            )}

            {/* Not loaded yet */}
            {!plData && !plLoading && (
              <div style={{ textAlign:"center", padding:"60px 20px", color:"#94a3b8" }}>
                <div style={{ fontSize:48 }}>🔬</div>
                <div style={{ fontWeight:700, color:"#374151", marginTop:12, fontSize:16 }}>Price Lab</div>
                <div style={{ fontSize:13, marginTop:6 }}>Load data first, then select a price type to simulate new prices</div>
              </div>
            )}

            {/* No price type selected */}
            {plData && !plabPriceType && (
              <div style={{ textAlign:"center", padding:"60px 20px", color:"#94a3b8" }}>
                <div style={{ fontSize:40 }}>🏷️</div>
                <div style={{ fontWeight:700, color:"#374151", marginTop:12 }}>Select a Price Type</div>
                <div style={{ fontSize:13, marginTop:6 }}>Choose a price type above to start simulating</div>
              </div>
            )}

            {plData && plabPriceType && (<>

              {/* KPI Bar */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:14 }}>
                {[
                  { label:"Total Items",    value:items.length,          color:"#0f2d5a",  icon:"📋" },
                  { label:"Simulated",      value:withNewPrice.length,   color:"#7c3aed",  icon:"✏️" },
                  { label:"Avg New Profit", value: avgNew !== null ? avgNew.toFixed(1)+"%" : "—", color: avgNew===null?"#94a3b8": avgNew>=40?"#15803d":avgNew>=20?"#d97706":"#dc2626", icon:"📈" },
                  { label:"🟢 Good",        value:green,                 color:"#15803d",  icon:"" },
                  { label:"🟡 Marginal",    value:yellow,                color:"#d97706",  icon:"" },
                  { label:"🔴 Low",         value:red,                   color:"#dc2626",  icon:"" },
                ].map(({label,value,color,icon})=>(
                  <div key={label} style={{ background:"#fff", borderRadius:10, padding:"12px 14px", border:"1px solid #e2e8f0", position:"relative", overflow:"hidden" }}>
                    <div style={{ position:"absolute", top:0, left:0, right:0, height:3, background:color }} />
                    <div style={{ fontSize:10, fontWeight:700, color:"#94a3b8", letterSpacing:1, textTransform:"uppercase", marginBottom:4 }}>{label}</div>
                    <div style={{ fontSize:20, fontWeight:900, fontFamily:"monospace", color }}>{value}</div>
                  </div>
                ))}
              </div>

              {/* Table + BOM Panel */}
              <div style={{ display:"grid", gridTemplateColumns: plabSelected ? "1fr 360px" : "1fr", gap:12, alignItems:"start" }}>
              <div style={{ overflowX:"auto", borderRadius:10, border:"1px solid #e2e8f0", boxShadow:"0 1px 4px rgba(0,0,0,0.06)" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, fontFamily:"inherit" }}>
                  <thead>
                    <tr style={{ background:"#f8fafc" }}>
                      <th style={{ padding:"10px 10px", textAlign:"left",  borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>#</th>
                      <th style={{ padding:"10px 6px",  textAlign:"center", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>BOM</th>
                      <th style={{ padding:"10px 10px", textAlign:"left",  borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Item Code</th>
                      <th style={{ padding:"10px 10px", textAlign:"left",  borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Description</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Conv.</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#0f2d5a", fontWeight:700 }}>Current SP</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Material Cost</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Total +35%</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Cur. Profit%</th>
                      <th style={{ padding:"10px 8px",  textAlign:"center",borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#7c3aed", fontWeight:700, background:"#f5f3ff", minWidth:120 }}>✏️ New Price</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#7c3aed", fontWeight:700, background:"#f5f3ff" }}>New Profit</th>
                      <th style={{ padding:"10px 8px",  textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#7c3aed", fontWeight:700, background:"#f5f3ff" }}>New Profit%</th>
                      {plabDiscount > 0 && <th style={{ padding:"10px 8px", textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#b45309", fontWeight:700, background:"#fffbeb", whiteSpace:"nowrap" }}>Net Price ({plabDiscount}% off)</th>}
                      {plabDiscount > 0 && <th style={{ padding:"10px 8px", textAlign:"right", borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#b45309", fontWeight:700, background:"#fffbeb" }}>Net Profit%</th>}
                      <th style={{ padding:"10px 8px",  textAlign:"center",borderBottom:"2px solid #e2e8f0", fontSize:11, color:"#475569", fontWeight:700 }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.length === 0 && (
                      <tr><td colSpan={12} style={{ padding:30, textAlign:"center", color:"#94a3b8" }}>No items found</td></tr>
                    )}
                    {items.map((r, i) => {
                      const locked = !r.hasCost;
                      const bg = locked ? "#fafafa" : i%2===0 ? "#fff" : "#fafafa";
                      return (
                        <tr key={r.code}
                          style={{ background: plabSelected===r.code ? "#eff6ff" : bg, borderBottom:"1px solid #f1f5f9", opacity: locked ? 0.55 : 1 }}>
                          <td style={{ padding:"8px 10px", color:"#94a3b8", fontSize:11 }}>{i+1}</td>
                          <td style={{ padding:"8px 6px", textAlign:"center" }}>
                            <button
                              onClick={e => { e.stopPropagation(); if (!locked) setPlabSelected(prev => prev === r.code ? null : r.code); }}
                              disabled={locked}
                              title={locked ? "No BOM" : "View BOM"}
                              style={{ background: plabSelected===r.code ? "#1d4ed8" : "#eff6ff", border:`1px solid ${plabSelected===r.code?"#1d4ed8":"#bfdbfe"}`, borderRadius:6, width:26, height:26, cursor: locked?"not-allowed":"pointer", fontSize:13, display:"inline-flex", alignItems:"center", justifyContent:"center", opacity: locked?0.35:1, transition:"all 0.15s" }}>
                              📋
                            </button>
                          </td>
                          <td style={{ padding:"8px 10px", fontWeight:700, fontFamily:"monospace", fontSize:11, color:"#1d4ed8" }}>{r.code}</td>
                          <td style={{ padding:"8px 10px", color:"#374151", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{r.desc}</td>
                          <td style={{ padding:"8px 8px", textAlign:"right", color:"#64748b", fontFamily:"monospace" }}>{r.conv ?? "—"}</td>
                          <td style={{ padding:"8px 8px", textAlign:"right", fontFamily:"monospace", fontWeight:700, color:"#0f2d5a" }}>{r.sp ? fmt(r.sp) : "—"}</td>
                          <td style={{ padding:"8px 8px", textAlign:"right", fontFamily:"monospace", color:"#475569" }}>{r.rawBoxCost ? fmt(r.rawBoxCost) : <span style={{color:"#dc2626"}}>No BOM</span>}</td>
                          <td style={{ padding:"8px 8px", textAlign:"right", fontFamily:"monospace", color:"#475569" }}>{r.boxCost ? fmt(r.boxCost) : "—"}</td>
                          <td style={{ padding:"8px 8px", textAlign:"right" }}>
                            {r.curProfitPct !== null
                              ? <span style={{ background:pctBg(r.curProfitPct), color:pctColor(r.curProfitPct), padding:"2px 7px", borderRadius:4, fontWeight:700, fontSize:11, fontFamily:"monospace" }}>{r.curProfitPct.toFixed(1)}%</span>
                              : <span style={{ color:"#94a3b8" }}>—</span>}
                          </td>
                          {/* New Price input */}
                          <td style={{ padding:"6px 8px", textAlign:"center", background:"#f5f3ff" }}>
                            {locked
                              ? <span style={{ fontSize:11, color:"#94a3b8" }}>🔒 No BOM</span>
                              : <input
                                  type="number" min="0" step="0.01"
                                  value={plabPrices[r.code] ?? 0}
                                  onClick={e => e.stopPropagation()}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setPlabPrices(prev => ({ ...prev, [r.code]: v }));
                                  }}
                                  style={{ width:"100%", padding:"5px 8px", border:"1px solid #d1d5db", borderRadius:6, fontSize:12, fontFamily:"monospace", textAlign:"right", outline:"none", background:"#fff", boxSizing:"border-box", MozAppearance:"textfield", WebkitAppearance:"none" }}
                                  onFocus={e=>e.target.style.borderColor="#6b7280"}
                                  onBlur={e=>e.target.style.borderColor="#d1d5db"}
                                />
                            }
                          </td>
                          {/* New Profit */}
                          <td style={{ padding:"8px 8px", textAlign:"right", fontFamily:"monospace", background:"#f5f3ff", fontWeight:700, color: r.newProfit===null?"#94a3b8": r.newProfit>=0?"#15803d":"#dc2626" }}>
                            {r.newProfit !== null ? fmt(r.newProfit) : "—"}
                          </td>
                          {/* New Profit % */}
                          <td style={{ padding:"8px 8px", textAlign:"right", background:"#f5f3ff" }}>
                            {r.newProfitPct !== null
                              ? <span style={{ background:pctBg(r.newProfitPct), color:pctColor(r.newProfitPct), padding:"2px 7px", borderRadius:4, fontWeight:700, fontSize:11, fontFamily:"monospace" }}>{r.newProfitPct.toFixed(1)}%</span>
                              : <span style={{ color:"#94a3b8" }}>—</span>}
                          </td>
                          {/* Net Price */}
                          {plabDiscount > 0 && (
                            <td style={{ padding:"8px 8px", textAlign:"right", fontFamily:"monospace", background:"#fffbeb", fontWeight:700, color: r.netPrice===null?"#94a3b8":"#b45309" }}>
                              {r.netPrice !== null ? fmt(r.netPrice) : locked ? "—" : "—"}
                            </td>
                          )}
                          {/* Net Profit % */}
                          {plabDiscount > 0 && (
                            <td style={{ padding:"8px 8px", textAlign:"right", background:"#fffbeb" }}>
                              {r.netProfitPct !== null
                                ? <span style={{ background:pctBg(r.netProfitPct), color:pctColor(r.netProfitPct), padding:"2px 7px", borderRadius:4, fontWeight:700, fontSize:11, fontFamily:"monospace" }}>{r.netProfitPct.toFixed(1)}%</span>
                                : <span style={{ color:"#94a3b8" }}>—</span>}
                            </td>
                          )}
                          {/* Status */}
                          <td style={{ padding:"8px 8px", textAlign:"center", fontSize:16 }}>
                            {r.newProfitPct !== null ? light(r.newProfitPct) : locked ? "🔒" : ""}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* BOM Detail Panel */}
              {plabSelected && (() => {
                const selR = items.find(r => r.code === plabSelected);
                if (!selR) return null;
                const bomItem = recalcData.find(r => r.code === plabSelected);
                return (
                  <div style={{ background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.08)", display:"flex", flexDirection:"column", maxHeight:"calc(100vh - 240px)", overflow:"hidden", position:"sticky", top:12 }}>
                    <div style={{ padding:"14px 18px", background:"linear-gradient(135deg, #0f2d5a, #1d4ed8)", color:"#fff" }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"start" }}>
                        <div>
                          <div style={{ fontWeight:800, fontSize:15, fontFamily:"monospace" }}>{selR.code}</div>
                          <div style={{ fontSize:12, opacity:0.8, marginTop:2, direction:"rtl" }}>{selR.desc}</div>
                        </div>
                        <button onClick={()=>setPlabSelected(null)} style={{ background:"rgba(255,255,255,0.2)", border:"none", color:"#fff", width:26, height:26, borderRadius:6, cursor:"pointer", fontSize:14 }}>✕</button>
                      </div>
                      {selR.hasCost && (
                        <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:6, marginTop:10 }}>
                          {[
                            { l:"Sell Price", v: selR.sp      ? fmt(selR.sp)      : "—" },
                            { l:"Material",   v: selR.rawBoxCost ? fmt(selR.rawBoxCost) : "—" },
                            { l:"+35% Total", v: selR.boxCost ? fmt(selR.boxCost) : "—" },
                            { l:"New Price",  v: selR.newSP   ? fmt(selR.newSP)   : "—" },
                          ].map((s,i) => (
                            <div key={i} style={{ background:"rgba(255,255,255,0.15)", borderRadius:7, padding:"7px 8px", textAlign:"center" }}>
                              <div style={{ fontSize:10, opacity:0.8 }}>{s.l}</div>
                              <div style={{ fontWeight:700, fontSize:13, marginTop:1, fontFamily:"monospace" }}>{s.v}</div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    <div style={{ overflowY:"auto", flex:1, padding:"14px 16px" }}>
                      {!selR.hasCost ? (
                        <div style={{ textAlign:"center", padding:"40px 16px", color:"#94a3b8" }}>
                          <div style={{ fontSize:40 }}>📭</div>
                          <div style={{ fontWeight:600, color:"#374151", marginTop:8 }}>No Bill of Materials</div>
                          <div style={{ fontSize:12, marginTop:4 }}>This item has no BOM entry</div>
                        </div>
                      ) : bomItem ? (
                        <div>
                          {bomItem.missing?.length > 0 && (
                            <div style={{ background:"#fffbeb", border:"1px solid #fcd34d", borderRadius:8, padding:"8px 12px", marginBottom:12, fontSize:12 }}>
                              <b style={{ color:"#92400e" }}>⚠️ Missing costs:</b> <span style={{ color:"#78350f" }}>{bomItem.missing.join(" • ")}</span>
                            </div>
                          )}
                          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                            <thead>
                              <tr style={{ background:"#f8fafc" }}>
                                <th style={{ padding:"7px 8px", textAlign:"left",  borderBottom:"2px solid #e2e8f0", color:"#475569" }}>Child Item</th>
                                <th style={{ padding:"7px 6px", textAlign:"right", borderBottom:"2px solid #e2e8f0", color:"#475569" }}>Qty</th>
                                <th style={{ padding:"7px 6px", textAlign:"right", borderBottom:"2px solid #e2e8f0", color:"#475569" }}>Unit Cost</th>
                                <th style={{ padding:"7px 6px", textAlign:"right", borderBottom:"2px solid #e2e8f0", color:"#475569" }}>Line Total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {bomItem.bomLines.map((l,i) => (
                                <tr key={i} style={{ borderBottom:"1px solid #f1f5f9", background: l.missing ? "#fef2f2" : "transparent" }}>
                                  <td style={{ padding:"6px 8px", fontWeight:600, color: l.missing?"#dc2626":"#0f2d5a", fontFamily:"monospace", fontSize:11 }}>{l.child}{l.missing && " ⚠️"}</td>
                                  <td style={{ padding:"6px", textAlign:"right", color:"#64748b", fontFamily:"monospace" }}>{fmtQty(l.qty)}</td>
                                  <td style={{ padding:"6px", textAlign:"right", fontFamily:"monospace" }}>{l.unitCost !== null ? fmt(l.unitCost) : <span style={{ color:"#dc2626" }}>—</span>}</td>
                                  <td style={{ padding:"6px", textAlign:"right", fontFamily:"monospace", fontWeight:600 }}>{fmt(l.lineCost)}</td>
                                </tr>
                              ))}
                            </tbody>
                            <tfoot>
                              <tr style={{ background:"#f8fafc", borderTop:"2px solid #e2e8f0" }}>
                                <td colSpan={3} style={{ padding:"8px", fontWeight:700 }}>Material Cost (Box)</td>
                                <td style={{ padding:"8px 6px", fontWeight:800, color:"#1d4ed8", fontFamily:"monospace", textAlign:"right" }}>{fmt(bomItem.rawBoxCost)}</td>
                              </tr>
                              <tr style={{ background:"#eff6ff" }}>
                                <td colSpan={3} style={{ padding:"8px", fontWeight:700, color:"#1e40af" }}>+ 35% Production</td>
                                <td style={{ padding:"8px 6px", fontWeight:800, color:"#1e40af", fontFamily:"monospace", textAlign:"right" }}>{fmt(bomItem.rawBoxCost * 0.35)}</td>
                              </tr>
                              <tr style={{ background:"#dbeafe" }}>
                                <td colSpan={3} style={{ padding:"8px", fontWeight:700, color:"#1d4ed8" }}>Total Cost (Box)</td>
                                <td style={{ padding:"8px 6px", fontWeight:800, color:"#1d4ed8", fontFamily:"monospace", textAlign:"right" }}>{fmt(bomItem.boxCost)}</td>
                              </tr>
                              {selR.newSP && (
                                <tr style={{ background:"#f5f3ff" }}>
                                  <td colSpan={3} style={{ padding:"8px", fontWeight:700, color:"#7c3aed" }}>New Price Profit</td>
                                  <td style={{ padding:"8px 6px", fontWeight:800, fontFamily:"monospace", textAlign:"right", color: selR.newProfit >= 0 ? "#15803d" : "#dc2626" }}>{fmt(selR.newProfit)}</td>
                                </tr>
                              )}
                              {selR.newSP && (
                                <tr style={{ background:"#f5f3ff" }}>
                                  <td colSpan={3} style={{ padding:"8px", fontWeight:700, color:"#7c3aed" }}>New Profit %</td>
                                  <td style={{ padding:"8px 6px", textAlign:"right" }}>
                                    <span style={{ background: selR.newProfitPct>=40?"#dcfce7":selR.newProfitPct>=20?"#fef3c7":"#fef2f2", color: selR.newProfitPct>=40?"#15803d":selR.newProfitPct>=20?"#d97706":"#dc2626", padding:"3px 8px", borderRadius:4, fontWeight:800, fontSize:12, fontFamily:"monospace" }}>{selR.newProfitPct?.toFixed(1)}%</span>
                                  </td>
                                </tr>
                              )}
                            </tfoot>
                          </table>
                        </div>
                      ) : (
                        <div style={{ textAlign:"center", padding:"40px 16px", color:"#94a3b8" }}>
                          <div style={{ fontSize:40 }}>📭</div>
                          <div style={{ fontWeight:600, color:"#374151", marginTop:8 }}>No BOM Data</div>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}
              </div>{/* end grid */}
            </>)}
          </div>
        );
        } catch(e) {
          return <div style={{padding:20,color:"#dc2626",fontWeight:700,fontSize:14}}>❌ Price Lab Error: {String(e.message)}</div>;
        }
      })()}


      {/* ── Customer Coverage ── */}
      {page === "coverage" && (() => {
        try {
        const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        const n = v => parseFloat(v) || 0;

        const coverageBar = (pct) => {
          const color = pct >= 60 ? "#16a34a" : pct >= 30 ? "#d97706" : "#dc2626";
          const bg    = pct >= 60 ? "#dcfce7" : pct >= 30 ? "#fef3c7" : "#fef2f2";
          return (
            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
              <div style={{ flex:1, height:8, background:"#e5e7eb", borderRadius:4, overflow:"hidden", minWidth:60 }}>
                <div style={{ height:"100%", width:Math.min(pct,100)+"%", background:color, borderRadius:4, transition:"width 0.3s" }} />
              </div>
              <span style={{ fontSize:11, fontWeight:700, color, fontFamily:"monospace", minWidth:38, textAlign:"right" }}>{pct.toFixed(1)}%</span>
            </div>
          );
        };

        const isCustomer = ccTab === "customer";
        const rawRows = isCustomer ? (ccData?.customers || []) : (ccData?.salespersons || []);

        // Filter by month + search (year handled by API per single year)
        const rows = rawRows.filter(r => {
          if (ccMonth > 0 && parseInt(r.InvoiceMonth) !== ccMonth) return false;
          if (ccSearch) {
            const s = ccSearch.toLowerCase();
            return (r.CustomerName||"").toLowerCase().includes(s) ||
                   (r.SalesName||"").toLowerCase().includes(s) ||
                   String(r.CustomerNo||"").includes(ccSearch);
          }
          return true;
        });

        // Sort
        const sorted = [...rows].sort((a,b) => {
          const d = ccSort.dir === "asc" ? 1 : -1;
          if (ccSort.col === "coverage")   return d * (n(b.FamilyCoveragePercent||b.AvgCustomerCoveragePercent) - n(a.FamilyCoveragePercent||a.AvgCustomerCoveragePercent));
          if (ccSort.col === "families")   return d * (n(b.ActiveFamilies||b.TotalCustomerActiveFamilies) - n(a.ActiveFamilies||a.TotalCustomerActiveFamilies));
          if (ccSort.col === "portfolio")  return d * (n(b.ActiveFamilies||b.TotalCustomerActiveFamilies) - n(a.ActiveFamilies||a.TotalCustomerActiveFamilies));
          if (ccSort.col === "amount")     return d * (n(b.TotalAmount) - n(a.TotalAmount));
          if (ccSort.col === "customers")  return d * (n(b.ActiveCustomers) - n(a.ActiveCustomers));
          if (ccSort.col === "name")      return d * (a.CustomerName||a.SalesName||"").localeCompare(b.CustomerName||b.SalesName||"");
          if (ccSort.col === "year")      return d * (n(a.InvoiceYear) - n(b.InvoiceYear));
          if (ccSort.col === "month")     return d * (n(a.InvoiceMonth) - n(b.InvoiceMonth));
          return 0;
        });

        // KPIs
        const avgCoverage = sorted.length ? sorted.reduce((s,r)=>s+n(isCustomer?r.FamilyCoveragePercent:r.AvgCustomerCoveragePercent),0)/sorted.length : 0;
        const totalAmt    = sorted.reduce((s,r)=>s+n(r.TotalAmount),0);
        const high  = sorted.filter(r=>n(isCustomer?r.FamilyCoveragePercent:r.AvgCustomerCoveragePercent)>=60).length;
        const mid   = sorted.filter(r=>{ const v=n(isCustomer?r.FamilyCoveragePercent:r.AvgCustomerCoveragePercent); return v>=30&&v<60; }).length;
        const low   = sorted.filter(r=>n(isCustomer?r.FamilyCoveragePercent:r.AvgCustomerCoveragePercent)<30).length;

        const SortTH = ({col,children,align="right"}) => {
          const active = ccSort.col===col;
          return (
            <th onClick={()=>setCcSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}
              style={{ padding:"10px 10px",textAlign:align,borderBottom:"2px solid #e2e8f0",fontSize:11,color:active?"#1d4ed8":"#475569",fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",userSelect:"none",background:"#f8fafc" }}>
              {children}{active?(ccSort.dir==="desc"?" ↓":" ↑"):" ↕"}
            </th>
          );
        };

        return (
          <div style={{ padding:"0 20px 20px" }}>

            {/* Toolbar */}
            <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap", marginBottom:14 }}>
              <button onClick={()=>fetchCoverage(ccYear)} disabled={ccLoading}
                style={{ background:ccLoading?"#94a3b8":"#1d4ed8",color:"#fff",border:"none",padding:"9px 18px",borderRadius:8,cursor:ccLoading?"not-allowed":"pointer",fontFamily:"inherit",fontSize:13,fontWeight:700 }}>
                {ccLoading?"⏳ Loading...":"🔄 Load Data"}
              </button>

              {/* Year */}
              <div style={{ display:"flex",alignItems:"center",gap:4,background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:8,padding:"4px 8px" }}>
                <button onClick={()=>{setCcYear(y=>y-1);}} style={{ background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b",padding:"0 4px" }}>‹</button>
                <span style={{ fontWeight:700,fontSize:14,minWidth:36,textAlign:"center" }}>{ccYear}</span>
                <button onClick={()=>{setCcYear(y=>y+1);}} style={{ background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b",padding:"0 4px" }}>›</button>
              </div>

              {/* Month filter */}
              <select value={ccMonth} onChange={e=>setCcMonth(parseInt(e.target.value))}
                style={{ padding:"8px 10px",border:"1px solid #d1d5db",borderRadius:8,fontSize:13,fontFamily:"inherit",background:"#fff",cursor:"pointer" }}>
                <option value={0}>All Months</option>
                {MONTHS.map((m,i)=><option key={i} value={i+1}>{m}</option>)}
              </select>

              {/* Search */}
              {ccData && (
                <input placeholder="🔍 Search..." value={ccSearch} onChange={e=>setCcSearch(e.target.value)}
                  style={{ padding:"8px 12px",border:"1px solid #d1d5db",borderRadius:8,fontSize:13,fontFamily:"inherit",width:180 }} />
              )}

              {ccMsg && <span style={{ fontSize:12,color:ccMsg.startsWith("✅")?"#15803d":"#dc2626",fontWeight:600 }}>{ccMsg}</span>}
            </div>

            {/* Tabs */}
            {ccData && (
              <div style={{ display:"flex",gap:4,marginBottom:14,background:"#f1f5f9",borderRadius:10,padding:4,width:"fit-content" }}>
                {[["customer","👤 By Customer"],["salesperson","🧑‍💼 By Salesperson"]].map(([k,lb])=>(
                  <button key={k} onClick={()=>{setCcTab(k);setCcSearch("");}}
                    style={{ padding:"8px 20px",borderRadius:8,border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:700,
                      background:ccTab===k?"#fff":"transparent",color:ccTab===k?"#1d4ed8":"#64748b",
                      boxShadow:ccTab===k?"0 1px 4px rgba(0,0,0,0.1)":"none",transition:"all 0.15s" }}>
                    {lb}
                  </button>
                ))}
              </div>
            )}

            {/* Empty state */}
            {!ccData && !ccLoading && (
              <div style={{ textAlign:"center",padding:"60px 20px",color:"#94a3b8" }}>
                <div style={{ fontSize:48 }}>📊</div>
                <div style={{ fontWeight:700,color:"#374151",marginTop:12,fontSize:16 }}>Customer Coverage</div>
                <div style={{ fontSize:13,marginTop:6 }}>Load data to see family coverage by customer and salesperson</div>
              </div>
            )}

            {ccData && (<>
              {/* KPI Bar */}
              <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:14 }}>
                {[
                  { label:"Records",       value:sorted.length,                 color:"#0f2d5a"  },
                  { label:"Avg Coverage",  value:avgCoverage.toFixed(1)+"%",    color: avgCoverage>=60?"#16a34a":avgCoverage>=30?"#d97706":"#dc2626" },
                  { label:"Total Amount",  value:fmt(totalAmt),                 color:"#0f2d5a"  },
                  { label:"🟢 High (≥60%)",value:high,                          color:"#16a34a"  },
                  { label:"🟡 Mid (30–60%)",value:mid,                          color:"#d97706"  },
                  { label:"🔴 Low (<30%)", value:low,                           color:"#dc2626"  },
                ].map(({label,value,color})=>(
                  <div key={label} style={{ background:"#fff",borderRadius:10,padding:"12px 14px",border:"1px solid #e2e8f0",position:"relative",overflow:"hidden" }}>
                    <div style={{ position:"absolute",top:0,left:0,right:0,height:3,background:color }} />
                    <div style={{ fontSize:10,fontWeight:700,color:"#94a3b8",letterSpacing:1,textTransform:"uppercase",marginBottom:4 }}>{label}</div>
                    <div style={{ fontSize:20,fontWeight:900,fontFamily:"monospace",color }}>{value}</div>
                  </div>
                ))}
              </div>

              {/* Table */}
              <div style={{ overflowX:"auto",borderRadius:10,border:"1px solid #e2e8f0",boxShadow:"0 1px 4px rgba(0,0,0,0.06)" }}>
                <table style={{ width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"inherit" }}>
                  <thead>
                    <tr>
                      <th style={{ padding:"10px 10px",textAlign:"left",borderBottom:"2px solid #e2e8f0",fontSize:11,color:"#475569",fontWeight:700,background:"#f8fafc" }}>#</th>
                      <SortTH col="month" align="left">Month</SortTH>
                      {isCustomer  && <SortTH col="name" align="left">👤 Customer</SortTH>}
                      {!isCustomer && <SortTH col="name" align="left">🧑‍💼 Salesperson</SortTH>}
                      {isCustomer  && <th style={{ padding:"10px 10px",textAlign:"left",borderBottom:"2px solid #e2e8f0",fontSize:11,color:"#475569",fontWeight:700,background:"#f8fafc" }}>Salesperson</th>}
                      {!isCustomer && <SortTH col="customers" align="right">Customers</SortTH>}
                      <SortTH col="families">Active Families</SortTH>
                      {!isCustomer && <SortTH col="avgfamilies" align="right">Avg Families</SortTH>}
                      <th style={{ padding:"10px 10px",textAlign:"right",borderBottom:"2px solid #e2e8f0",fontSize:11,color:"#475569",fontWeight:700,background:"#f8fafc" }}>Total Families</th>
                      <SortTH col="coverage">Avg Coverage %</SortTH>
                      <SortTH col="portfolio" align="right">Portfolio</SortTH>
                      <SortTH col="amount">Amount</SortTH>
                    </tr>
                  </thead>
                  <tbody>
                    {sorted.length === 0 && (
                      <tr><td colSpan={9} style={{ padding:30,textAlign:"center",color:"#94a3b8" }}>No data found</td></tr>
                    )}
                    {sorted.map((r,i) => {
                      const pct = n(isCustomer ? r.FamilyCoveragePercent : r.AvgCustomerCoveragePercent);
                      const bg  = i%2===0?"#fff":"#f8fafc";
                      return (
                        <tr key={i} style={{ background:bg,borderBottom:"1px solid #f1f5f9" }}>
                          <td style={{ padding:"8px 10px",color:"#94a3b8",fontSize:11 }}>{i+1}</td>
                          <td style={{ padding:"8px 10px",fontWeight:600,color:"#475569" }}>{MONTHS[(parseInt(r.InvoiceMonth)||1)-1]}</td>
                          {isCustomer && (
                            <td style={{ padding:"8px 10px" }}>
                              <div style={{ fontWeight:700,fontSize:12,color:"#0f2d5a" }}>{r.CustomerName||"—"}</div>
                              <div style={{ fontSize:10,color:"#94a3b8",fontFamily:"monospace" }}>{r.CustomerNo}</div>
                            </td>
                          )}
                          {!isCustomer && (
                            <td style={{ padding:"8px 10px",fontWeight:700,color:"#0f2d5a" }}>{r.SalesName||"—"}</td>
                          )}
                          {isCustomer  && <td style={{ padding:"8px 10px",color:"#64748b",fontSize:11 }}>{r.SalesName||"—"}</td>}
                          {!isCustomer && <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#1d4ed8" }}>{r.ActiveCustomers||"—"}</td>}
                          <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#374151" }}>{isCustomer ? r.ActiveFamilies : r.TotalCustomerActiveFamilies}</td>
                          {!isCustomer && <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",color:"#64748b" }}>{r.AvgFamiliesPerCustomer||"—"}</td>}
                          <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",color:"#94a3b8" }}>{r.TotalFamilies}</td>
                          <td style={{ padding:"8px 16px",minWidth:160 }}>{coverageBar(n(isCustomer ? r.FamilyCoveragePercent : r.AvgCustomerCoveragePercent))}</td>
                          <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#374151" }}>
                            {r.TotalFamilies ? `${r.ActiveFamilies||r.TotalCustomerActiveFamilies||0}/${r.TotalFamilies}` : "—"}
                          </td>
                          <td style={{ padding:"8px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#0f2d5a" }}>{fmt(n(r.TotalAmount))}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  {sorted.length > 0 && (
                    <tfoot>
                      <tr style={{ background:"#f1f5f9",borderTop:"2px solid #e2e8f0" }}>
                        <td colSpan={isCustomer?6:7} style={{ padding:"9px 10px",fontWeight:700,fontSize:12 }}>TOTAL / AVG</td>
                        <td style={{ padding:"9px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#1d4ed8" }}>{avgCoverage.toFixed(1)}%</td>
                        <td />{/* portfolio */}
                        <td style={{ padding:"9px 10px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:"#0f2d5a" }}>{fmt(totalAmt)}</td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </>)}
          </div>
        );
        } catch(e) {
          return <div style={{padding:20,color:"#dc2626",fontWeight:700,fontSize:14}}>❌ Coverage Error: {String(e.message)}</div>;
        }
      })()}

      {/* ── Expenses Analysis Page ── */}
      {page === "expenses" && (() => {
        try {
          const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

          const allParents = [...new Set(exData.map(r => r.Parent).filter(Boolean))].sort();
          const filteredParents = allParents.filter(p =>
            !exParentSearch || p.toLowerCase().includes(exParentSearch.toLowerCase())
          );

          const filtered = exData.filter(r => {
            const monthOk  = !exMonth.length  || exMonth.includes(parseInt(r.JournalMonth));
            const parentOk = !exParent.length || exParent.includes(String(r.Parent||""));
            const q = exSearch.toLowerCase();
            const searchOk = !q
              || String(r.Segment9||"").toLowerCase().includes(q)
              || (r.ValueDescription||"").toLowerCase().includes(q)
              || (r.Parent||"").toLowerCase().includes(q);
            return monthOk && parentOk && searchOk;
          });

          const groupKey = row => {
            const parts = [];
            if (exGroupBy.includes("account")) parts.push(String(row.Segment9||"") + "|" + (row.ValueDescription||""));
            if (exGroupBy.includes("parent"))  parts.push(row.Parent||"(No Parent)");
            if (exGroupBy.includes("month"))   parts.push(String(parseInt(row.JournalMonth)||""));
            return parts.join("||") || "All";
          };

          const getAmt = r => parseFloat(r.Amount ?? r.amount ?? r.NetAmount ?? r.netAmount ?? 0) || 0;

          const aggMap = {};
          filtered.forEach(r => {
            const k = groupKey(r);
            if (!aggMap[k]) aggMap[k] = {
              _key: k,
              Segment9: r.Segment9,
              ValueDescription: r.ValueDescription,
              Parent: r.Parent || "(No Parent)",
              JournalMonth: r.JournalMonth,
              Amount: 0,
              _rows: [],
            };
            aggMap[k].Amount += getAmt(r);
            aggMap[k]._rows.push(r);
          });

          let rows = Object.values(aggMap);

          rows.sort((a, b) => {
            const d = exSort.dir === "asc" ? 1 : -1;
            if (exSort.col === "Amount")           return d * (a.Amount - b.Amount);
            if (exSort.col === "Segment9")         return d * String(a.Segment9||"").localeCompare(String(b.Segment9||""));
            if (exSort.col === "ValueDescription") return d * (a.ValueDescription||"").localeCompare(b.ValueDescription||"");
            if (exSort.col === "Parent")           return d * (a.Parent||"").localeCompare(b.Parent||"");
            if (exSort.col === "Month")            return d * ((parseInt(a.JournalMonth)||0) - (parseInt(b.JournalMonth)||0));
            return 0;
          });

          const grandTotal = rows.reduce((s, r) => s + r.Amount, 0);

          const amounts = [...rows].map(r => r.Amount).sort((a,b)=>b-a);
          const p25 = amounts[Math.floor(amounts.length * 0.25)] ?? 0;
          const p75 = amounts[Math.floor(amounts.length * 0.75)] ?? 0;
          const light = amt => amt >= p25 ? "red" : amt >= p75 ? "orange" : "green";
          const lightColor = l => l==="red"?"#ef4444":l==="orange"?"#f59e0b":"#22c55e";

          const kpiRecords  = rows.length;
          const kpiTotal    = grandTotal;
          const kpiAvgMonth = (() => {
            const byMonth = {};
            filtered.forEach(r => { byMonth[parseInt(r.JournalMonth)] = (byMonth[parseInt(r.JournalMonth)]||0) + getAmt(r); });
            const vals = Object.values(byMonth);
            return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : 0;
          })();
          const kpiHigh = rows.filter(r => light(r.Amount) === "red").length;
          const kpiMed  = rows.filter(r => light(r.Amount) === "orange").length;
          const kpiLow  = rows.filter(r => light(r.Amount) === "green").length;

          const fmt  = n => (parseFloat(n)||0).toLocaleString("en-US", { minimumFractionDigits:0, maximumFractionDigits:0 });
          const fmtK = n => { const v=parseFloat(n)||0; return Math.abs(v)>=1e6?(v/1e6).toFixed(1)+"M":Math.abs(v)>=1e3?(v/1e3).toFixed(1)+"K":fmt(v); };

          const selectedRow = exSelected ? rows.find(r => r._key === exSelected) : null;

          const detailMonthly = selectedRow ? (() => {
            const map = {};
            selectedRow._rows.forEach(r => { const m=parseInt(r.JournalMonth); map[m]=(map[m]||0)+getAmt(r); });
            const total = Object.values(map).reduce((s,v)=>s+v,0);
            return MONTHS.map((name,i)=>({ month:name, num:i+1, amount:map[i+1]||0, share: total?((map[i+1]||0)/total*100).toFixed(1):"0.0" })).filter(x=>x.amount);
          })() : [];

          const detailSub = selectedRow ? (() => {
            const subMap = {};
            selectedRow._rows.forEach(r => {
              const k = exGroupBy.includes("parent")
                ? (String(r.Segment9||"") + " – " + (r.ValueDescription||""))
                : (r.Parent||"(No Parent)");
              subMap[k] = (subMap[k]||0) + getAmt(r);
            });
            const total = Object.values(subMap).reduce((s,v)=>s+v,0);
            return Object.entries(subMap)
              .map(([name,amount])=>({ name, amount, share: total?(amount/total*100).toFixed(1):"0.0" }))
              .sort((a,b)=>b.amount-a.amount);
          })() : [];

          const sortCol  = col => setExSort(s => ({ col, dir: s.col===col && s.dir==="desc" ? "asc" : "desc" }));
          const sortArrow = col => exSort.col===col ? (exSort.dir==="desc"?"↓":"↑") : "↕";
          const toggleGroup = val => setExGroupBy(prev =>
            prev.includes(val) ? (prev.length > 1 ? prev.filter(x=>x!==val) : prev) : [...prev, val]
          );

          const TH = ({ col, children, align="left" }) => (
            <th onClick={()=>sortCol(col)}
              style={{ padding:"10px 12px", textAlign:align, cursor:"pointer", whiteSpace:"nowrap",
                color: exSort.col===col?"#7c3aed":"#64748b", fontWeight:600, fontSize:12,
                borderBottom:"2px solid #e2e8f0", background:"#f8fafc", userSelect:"none" }}>
              {children} {sortArrow(col)}
            </th>
          );

          return (
            <div style={{ padding:"24px 28px", fontFamily:"inherit" }}
              onClick={()=>{ if(exMonthOpen||exParentOpen||exGroupOpen){ setExMonthOpen(false); setExParentOpen(false); setExGroupOpen(false); } }}
            >
              {/* Header */}
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20 }}>
                <div>
                  <h2 style={{ margin:0, fontSize:22, fontWeight:700, color:"#1e293b", display:"flex", alignItems:"center", gap:8 }}>
                    💸 Expenses Analysis
                  </h2>
                  <div style={{ fontSize:13, color:"#64748b", marginTop:3 }}>Account 6011 · General Ledger · {exYear}</div>
                </div>
                <div style={{ display:"flex", alignItems:"center", gap:6, background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:10, padding:"6px 14px" }}>
                  <button onClick={e=>{e.stopPropagation();setExYear(y=>y-1);setExSelected(null);}}
                    style={{ border:"none", background:"none", cursor:"pointer", fontSize:18, color:"#7c3aed", fontWeight:700, padding:"0 4px" }}>‹</button>
                  <span style={{ fontWeight:700, fontSize:16, minWidth:44, textAlign:"center" }}>{exYear}</span>
                  <button onClick={e=>{e.stopPropagation();setExYear(y=>y+1);setExSelected(null);}}
                    style={{ border:"none", background:"none", cursor:"pointer", fontSize:18, color:"#7c3aed", fontWeight:700, padding:"0 4px" }}>›</button>
                </div>
              </div>

              {/* Toolbar */}
              <div style={{ display:"flex", flexWrap:"wrap", gap:10, marginBottom:18, alignItems:"center" }}>
                {/* Load Data */}
                <button onClick={e=>{e.stopPropagation();fetchExpenses(exYear);}} disabled={exLoading}
                  style={{ background:exLoading?"#a78bfa":"#7c3aed", color:"#fff", border:"none", borderRadius:8,
                    padding:"8px 20px", fontWeight:600, fontSize:14, cursor:exLoading?"not-allowed":"pointer" }}>
                  {exLoading ? "⏳ Loading…" : "⚡ Load Data"}
                </button>

                {/* Month filter */}
                <div style={{ position:"relative" }} onClick={e=>e.stopPropagation()}>
                  <button onClick={()=>{ setExMonthOpen(o=>!o); setExParentOpen(false); setExGroupOpen(false); }}
                    style={{ background:"#f1f5f9", border:"1px solid #e2e8f0", borderRadius:8, padding:"7px 14px", cursor:"pointer", fontSize:13, display:"flex", alignItems:"center", gap:6 }}>
                    📅 Month {exMonth.length ? `(${exMonth.length})` : "(All)"} ▾
                  </button>
                  {exMonthOpen && (
                    <div style={{ position:"absolute", top:"110%", left:0, zIndex:200, background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, boxShadow:"0 8px 24px #0002", minWidth:160, padding:8 }}>
                      <div style={{ display:"flex", gap:6, padding:"4px 6px", borderBottom:"1px solid #f1f5f9", marginBottom:4 }}>
                        <span onClick={()=>setExMonth([])} style={{ fontSize:12, color:"#7c3aed", cursor:"pointer", fontWeight:600 }}>All</span>
                        <span style={{ color:"#cbd5e1" }}>|</span>
                        <span onClick={()=>setExMonth(MONTHS.map((_,i)=>i+1))} style={{ fontSize:12, color:"#7c3aed", cursor:"pointer", fontWeight:600 }}>Select All</span>
                      </div>
                      {MONTHS.map((m,i)=>(
                        <div key={i} onClick={()=>setExMonth(prev=>prev.includes(i+1)?prev.filter(x=>x!==i+1):[...prev,i+1])}
                          style={{ padding:"6px 10px", cursor:"pointer", display:"flex", alignItems:"center", gap:8, borderRadius:6,
                            background:exMonth.includes(i+1)?"#f3f0ff":"transparent" }}>
                          <span style={{ width:16,height:16,borderRadius:4,border:"2px solid #7c3aed",display:"flex",alignItems:"center",justifyContent:"center",background:exMonth.includes(i+1)?"#7c3aed":"transparent" }}>
                            {exMonth.includes(i+1)&&<span style={{ color:"#fff",fontSize:10,fontWeight:700 }}>✓</span>}
                          </span>
                          <span style={{ fontSize:13 }}>{m}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Parent filter */}
                <div style={{ position:"relative" }} onClick={e=>e.stopPropagation()}>
                  <button onClick={()=>{ setExParentOpen(o=>!o); setExMonthOpen(false); setExGroupOpen(false); }}
                    style={{ background:"#f1f5f9", border:"1px solid #e2e8f0", borderRadius:8, padding:"7px 14px", cursor:"pointer", fontSize:13, display:"flex", alignItems:"center", gap:6 }}>
                    🏷️ Parent {exParent.length ? `(${exParent.length})` : "(All)"} ▾
                  </button>
                  {exParentOpen && (
                    <div style={{ position:"absolute", top:"110%", left:0, zIndex:200, background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, boxShadow:"0 8px 24px #0002", minWidth:220, padding:8 }}>
                      <input value={exParentSearch} onChange={e=>setExParentSearch(e.target.value)}
                        placeholder="🔍 Search parent…"
                        style={{ width:"100%", border:"1px solid #e2e8f0", borderRadius:6, padding:"5px 8px", fontSize:12, marginBottom:6, boxSizing:"border-box" }} />
                      <div style={{ display:"flex", gap:6, padding:"4px 6px", borderBottom:"1px solid #f1f5f9", marginBottom:4 }}>
                        <span onClick={()=>{ setExParent([]); setExParentSearch(""); }} style={{ fontSize:12, color:"#7c3aed", cursor:"pointer", fontWeight:600 }}>All</span>
                        <span style={{ color:"#cbd5e1" }}>|</span>
                        <span onClick={()=>{ setExParent(filteredParents); setExParentSearch(""); }} style={{ fontSize:12, color:"#7c3aed", cursor:"pointer", fontWeight:600 }}>Select All</span>
                      </div>
                      <div style={{ maxHeight:200, overflowY:"auto" }}>
                        {filteredParents.map(p=>(
                          <div key={p} onClick={()=>setExParent(prev=>prev.includes(p)?prev.filter(x=>x!==p):[...prev,p])}
                            style={{ padding:"6px 10px", cursor:"pointer", display:"flex", alignItems:"center", gap:8, borderRadius:6,
                              background:exParent.includes(p)?"#f3f0ff":"transparent" }}>
                            <span style={{ width:16,height:16,borderRadius:4,border:"2px solid #7c3aed",display:"flex",alignItems:"center",justifyContent:"center",background:exParent.includes(p)?"#7c3aed":"transparent",flexShrink:0 }}>
                              {exParent.includes(p)&&<span style={{ color:"#fff",fontSize:10,fontWeight:700 }}>✓</span>}
                            </span>
                            <span style={{ fontSize:13 }}>{p}</span>
                          </div>
                        ))}
                        {!filteredParents.length && <div style={{ fontSize:12, color:"#94a3b8", padding:"6px 10px" }}>No results</div>}
                      </div>
                    </div>
                  )}
                </div>

                {/* Group By */}
                <div style={{ position:"relative" }} onClick={e=>e.stopPropagation()}>
                  <button onClick={()=>{ setExGroupOpen(o=>!o); setExMonthOpen(false); setExParentOpen(false); }}
                    style={{ background:"#f1f5f9", border:"1px solid #e2e8f0", borderRadius:8, padding:"7px 14px", cursor:"pointer", fontSize:13, display:"flex", alignItems:"center", gap:6 }}>
                    📊 Group By ({exGroupBy.length}) ▾
                  </button>
                  {exGroupOpen && (
                    <div style={{ position:"absolute", top:"110%", left:0, zIndex:200, background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, boxShadow:"0 8px 24px #0002", minWidth:180, padding:8 }}>
                      {[["account","Per Account"],["parent","Per Parent"],["month","Per Month"]].map(([val,label])=>(
                        <div key={val} onClick={()=>toggleGroup(val)}
                          style={{ padding:"7px 10px", cursor:"pointer", display:"flex", alignItems:"center", gap:8, borderRadius:6,
                            background:exGroupBy.includes(val)?"#f3f0ff":"transparent" }}>
                          <span style={{ width:16,height:16,borderRadius:4,border:"2px solid #7c3aed",display:"flex",alignItems:"center",justifyContent:"center",background:exGroupBy.includes(val)?"#7c3aed":"transparent" }}>
                            {exGroupBy.includes(val)&&<span style={{ color:"#fff",fontSize:10,fontWeight:700 }}>✓</span>}
                          </span>
                          <span style={{ fontSize:13 }}>{label}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Search */}
                <input value={exSearch} onChange={e=>{setExSearch(e.target.value);setExSelected(null);}}
                  onClick={e=>e.stopPropagation()}
                  placeholder="🔍 Search code, name, parent…"
                  style={{ border:"1px solid #e2e8f0", borderRadius:8, padding:"7px 14px", fontSize:13, minWidth:240, outline:"none" }} />

                {exMsg && <span style={{ fontSize:13, color:exMsg.startsWith("✅")?"#15803d":"#ef4444" }}>{exMsg}</span>}
              </div>

              {/* KPI Cards */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:12, marginBottom:20 }}>
                {[
                  { label:"Records",      value:kpiRecords,      accent:"#6366f1" },
                  { label:"Total Amount", value:fmtK(kpiTotal),  accent:"#0ea5e9" },
                  { label:"Avg / Month",  value:fmtK(kpiAvgMonth), accent:"#8b5cf6" },
                  { label:"🔴 High",      value:kpiHigh,         accent:"#ef4444" },
                  { label:"🟡 Medium",    value:kpiMed,          accent:"#f59e0b" },
                  { label:"🟢 Low",       value:kpiLow,          accent:"#22c55e" },
                ].map(({label,value,accent})=>(
                  <div key={label} style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:12, padding:"14px 16px",
                    borderTop:`3px solid ${accent}`, boxShadow:"0 1px 4px #0001" }}>
                    <div style={{ fontSize:11, color:"#94a3b8", marginBottom:4 }}>{label}</div>
                    <div style={{ fontSize:20, fontWeight:700, color:"#1e293b" }}>{value}</div>
                  </div>
                ))}
              </div>

              {/* Table + Detail Panel */}
              <div style={{ display:"flex", gap:16, alignItems:"flex-start" }}>

                {/* Table */}
                <div style={{ flex:1, overflowX:"auto", background:"#fff", border:"1px solid #e2e8f0", borderRadius:12, boxShadow:"0 1px 4px #0001" }}>
                  {exLoading
                    ? <div style={{ padding:40, textAlign:"center", color:"#94a3b8" }}>⏳ Loading…</div>
                    : !rows.length
                      ? <div style={{ padding:40, textAlign:"center", color:"#94a3b8" }}>
                          {exData.length ? "No rows match filters." : "Click ⚡ Load Data to begin."}
                        </div>
                      : (
                        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                          <thead>
                            <tr>
                              <th style={{ padding:"10px 12px", textAlign:"center", borderBottom:"2px solid #e2e8f0", background:"#f8fafc", color:"#94a3b8", fontSize:12, fontWeight:600 }}>#</th>
                              {exGroupBy.includes("account") && <>
                                <TH col="Segment9" align="left">Code</TH>
                                <TH col="ValueDescription" align="left">Account Name</TH>
                              </>}
                              {exGroupBy.includes("parent") && <TH col="Parent" align="left">Parent</TH>}
                              {exGroupBy.includes("month")  && <TH col="Month" align="left">Month</TH>}
                              <TH col="Amount" align="right">Amount</TH>
                              <th style={{ padding:"10px 12px", borderBottom:"2px solid #e2e8f0", background:"#f8fafc", color:"#64748b", fontWeight:600, fontSize:12, minWidth:140 }}>Share %</th>
                            </tr>
                          </thead>
                          <tbody>
                            {rows.map((row,idx)=>{
                              const l = light(row.Amount);
                              const share = grandTotal ? (row.Amount/grandTotal*100) : 0;
                              const isSelected = exSelected === row._key;
                              return (
                                <tr key={row._key}
                                  onClick={()=>{ setExSelected(isSelected?null:row._key); setExDetailTab("monthly"); }}
                                  style={{ background:isSelected?"#f3f0ff":idx%2===0?"#fff":"#fafafa",
                                    borderBottom:"1px solid #f1f5f9", cursor:"pointer", transition:"background 0.12s" }}>
                                  <td style={{ padding:"9px 12px", color:"#94a3b8", fontSize:11, textAlign:"center" }}>{idx+1}</td>
                                  {exGroupBy.includes("account") && <>
                                    <td style={{ padding:"9px 12px", fontFamily:"monospace", color:"#475569", fontWeight:600 }}>{row.Segment9}</td>
                                    <td style={{ padding:"9px 12px", color:"#1e293b" }}>{row.ValueDescription}</td>
                                  </>}
                                  {exGroupBy.includes("parent") && <td style={{ padding:"9px 12px", color:"#7c3aed", fontWeight:600 }}>{row.Parent}</td>}
                                  {exGroupBy.includes("month")  && <td style={{ padding:"9px 12px", color:"#64748b" }}>{MONTHS[(parseInt(row.JournalMonth)||1)-1]}</td>}
                                  <td style={{ padding:"9px 12px", textAlign:"right", fontWeight:700, color:lightColor(l), fontFamily:"monospace" }}>{fmt(row.Amount)}</td>
                                  <td style={{ padding:"9px 12px" }}>
                                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                      <div style={{ flex:1, height:6, background:"#f1f5f9", borderRadius:3 }}>
                                        <div style={{ width:`${Math.min(share,100)}%`, height:"100%", background:lightColor(l), borderRadius:3 }} />
                                      </div>
                                      <span style={{ fontSize:11, color:"#64748b", minWidth:36, textAlign:"right" }}>{share.toFixed(1)}%</span>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot>
                            <tr style={{ background:"#f8fafc", borderTop:"2px solid #e2e8f0", fontWeight:700 }}>
                              <td style={{ padding:"9px 12px", color:"#64748b", fontSize:12 }}
                                colSpan={1 + (exGroupBy.includes("account")?2:0) + (exGroupBy.includes("parent")?1:0) + (exGroupBy.includes("month")?1:0)}>
                                TOTAL ({rows.length} rows)
                              </td>
                              <td style={{ padding:"9px 12px", textAlign:"right", color:"#1e293b", fontFamily:"monospace" }}>{fmt(grandTotal)}</td>
                              <td style={{ padding:"9px 12px", color:"#64748b", fontSize:12 }}>100.0%</td>
                            </tr>
                          </tfoot>
                        </table>
                      )
                  }
                </div>

                {/* Detail Panel */}
                {selectedRow && (
                  <div style={{ width:320, background:"#fff", border:"1px solid #e2e8f0", borderRadius:12, boxShadow:"0 4px 16px #0002", flexShrink:0 }}>
                    <div style={{ padding:"14px 16px", borderBottom:"1px solid #f1f5f9", display:"flex", alignItems:"flex-start", justifyContent:"space-between" }}>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ fontWeight:700, fontSize:14, color:"#1e293b", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {exGroupBy.includes("account") ? (selectedRow.ValueDescription || selectedRow.Segment9) : selectedRow.Parent}
                        </div>
                        {exGroupBy.includes("account") && <div style={{ fontSize:11, color:"#94a3b8", marginTop:2 }}>Code: {selectedRow.Segment9}</div>}
                        <div style={{ fontSize:13, fontWeight:700, color:"#7c3aed", marginTop:4, fontFamily:"monospace" }}>{fmt(selectedRow.Amount)}</div>
                      </div>
                      <button onClick={()=>setExSelected(null)}
                        style={{ border:"none", background:"#f1f5f9", borderRadius:6, width:26, height:26, cursor:"pointer", fontSize:14, color:"#64748b", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>✕</button>
                    </div>
                    <div style={{ display:"flex", borderBottom:"1px solid #f1f5f9" }}>
                      {[["monthly","📅 Monthly"],["sub","🏷️ Sub-accounts"]].map(([t,label])=>(
                        <button key={t} onClick={()=>setExDetailTab(t)}
                          style={{ flex:1, padding:"9px 0", border:"none", background:"none", cursor:"pointer", fontSize:12,
                            fontWeight:exDetailTab===t?700:400, color:exDetailTab===t?"#7c3aed":"#94a3b8",
                            borderBottom:exDetailTab===t?"2px solid #7c3aed":"2px solid transparent" }}>{label}</button>
                      ))}
                    </div>
                    <div style={{ padding:12, maxHeight:400, overflowY:"auto" }}>
                      {exDetailTab === "monthly" && (
                        detailMonthly.length
                          ? detailMonthly.map(({month,amount,share})=>(
                            <div key={month} style={{ marginBottom:10 }}>
                              <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3 }}>
                                <span style={{ fontWeight:600, color:"#475569" }}>{month}</span>
                                <span style={{ color:"#1e293b", fontWeight:700, fontFamily:"monospace" }}>{fmt(amount)}</span>
                              </div>
                              <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                <div style={{ flex:1, height:5, background:"#f1f5f9", borderRadius:3 }}>
                                  <div style={{ width:`${share}%`, height:"100%", background:"#7c3aed", borderRadius:3 }} />
                                </div>
                                <span style={{ fontSize:11, color:"#94a3b8", minWidth:32, textAlign:"right" }}>{share}%</span>
                              </div>
                            </div>
                          ))
                          : <div style={{ fontSize:12, color:"#94a3b8", textAlign:"center", padding:20 }}>No monthly data</div>
                      )}
                      {exDetailTab === "sub" && (
                        detailSub.length
                          ? detailSub.map(({name,amount,share})=>(
                            <div key={name} style={{ marginBottom:10 }}>
                              <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3 }}>
                                <span style={{ fontWeight:600, color:"#475569", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{name}</span>
                                <span style={{ color:"#1e293b", fontWeight:700, fontFamily:"monospace" }}>{fmt(amount)}</span>
                              </div>
                              <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                <div style={{ flex:1, height:5, background:"#f1f5f9", borderRadius:3 }}>
                                  <div style={{ width:`${share}%`, height:"100%", background:"#0ea5e9", borderRadius:3 }} />
                                </div>
                                <span style={{ fontSize:11, color:"#94a3b8", minWidth:32, textAlign:"right" }}>{share}%</span>
                              </div>
                            </div>
                          ))
                          : <div style={{ fontSize:12, color:"#94a3b8", textAlign:"center", padding:20 }}>No sub-account data</div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          );
        } catch(e) {
          return <div style={{ padding:20, color:"#dc2626", fontWeight:700, fontSize:14 }}>❌ Expenses Error: {String(e.message)}</div>;
        }
      })()}

      </div>{/* end flex:1 page content */}
      </div>{/* end cm-main */}
    </div>}
    </>
  );
}