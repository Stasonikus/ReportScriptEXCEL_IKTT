import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(root, "docs", "Power_BI_technical_specification.xlsx");
const previewDir = path.join(root, "tmp", "power_bi_spec_previews_v2");
const pointsPath = path.join(root, "src", "data", "border_crossing_points.json");
const wb = Workbook.create();
const NAVY = "#17365D", PALE = "#F3F7FA", AMBER = "#FFF2CC", TEXT = "#1F2937", FONT = "Arial";

function colName(index) { let n = index + 1, text = ""; while (n > 0) { n--; text = String.fromCharCode(65 + (n % 26)) + text; n = Math.floor(n / 26); } return text; }
function addSheet(name, title, description, headers, rows, tableName, widths, tabColor) {
  const sh = wb.worksheets.add(name); sh.showGridLines = false; if (tabColor) sh.tabColor = tabColor;
  sh.getRange("A2").values = [[title]]; sh.getRange("A2").format.font = {name: FONT, size: 14, bold: true, color: "#000000"};
  sh.getRange("A3").values = [[description]]; sh.getRange("A3").format.font = {name: FONT, size: 10, italic: true, color: "#595959"};
  const start = 5, end = start + rows.length, last = colName(headers.length - 1);
  const range = sh.getRange(`A${start}:${last}${end}`); range.values = [headers, ...rows];
  range.format.font = {name: FONT, size: 10, color: TEXT}; range.format.wrapText = true; range.format.verticalAlignment = "center"; range.format.borders = {preset:"all",style:"thin",color:"#D9D9D9"};
  const header = sh.getRange(`A${start}:${last}${start}`); header.format.fill = NAVY; header.format.font = {name:FONT,size:10,bold:true,color:"#FFFFFF"}; header.format.horizontalAlignment = "center"; header.format.rowHeight = 30;
  for (let r = start + 1; r <= end; r++) if ((r - start) % 2 === 0) sh.getRange(`A${r}:${last}${r}`).format.fill = PALE;
  const table = sh.tables.add(`A${start}:${last}${end}`, true, tableName); table.style = "TableStyleMedium2"; table.showFilterButton = true;
  sh.freezePanes.freezeRows(start); headers.forEach((_, i) => sh.getRange(`${colName(i)}1:${colName(i)}${Math.max(25,end)}`).format.columnWidth = widths[i] || 20);
  sh.getRange(`A${start + 1}:${last}${end}`).format.autofitRows(); return {sh, end, last};
}

addSheet("README", "Техническая спецификация данных для Power BI", "Один запуск использует три логических входа и создаёт один Excel набор с постоянной схемой.", ["Параметр","Значение"], [
  ["Назначение","Набор данных для Power BI по перевозкам, пунктам пропуска и навигационным пломбам"],
  ["Начало второго этапа","31.07.2026 00:00 UTC+05:00"], ["Версия схемы","2.0"],
  ["Входы","Большая выгрузка, реестр пунктов с координатами, единый снимок пломб"],
  ["Транспорт","AUTO и RAIL определяются по большой общей выгрузке"],
  ["Снимок пломб","Одна пломба на дату среза, с владельцем, местом, координатами и качеством"],
  ["Выход","out/power bi data/power_bi_dataset_<timestamp>.xlsx"],
  ["Правило пустых значений","Пусто означает отсутствуют данные; 0 означает рассчитанное отсутствие; UNKNOWN означает нераспознанное значение"],
  ["Основной документ","Power_BI_technical_specification.docx"]
], "ReadmeTable", [30,110], NAVY);

addSheet("ai_handoff", "Порядок работы для разработчика или ИИ", "Эта последовательность обязательна для каждого обновления.", ["Шаг","Действие","Требование"], [
  [1,"Получить входы","Принять только большую выгрузку, реестр пунктов и снимок пломб. Имена файлов могут отличаться, структура заголовков обязательна."],
  [2,"Проверить входы","Записать имя файла, дату изменения, контрольную сумму и отсутствующие столбцы в dim_source и dq_issues."],
  [3,"Ограничить второй этап","Оставить created_at не раньше 31.07.2026 00:00 UTC+05:00."],
  [4,"Нормализовать перевозки","Выбрать последнюю версию transportation_id по status_changed_at; сохранить исходные значения и quality_code."],
  [5,"Определить транспорт","Использовать Тип транспорта, затем номера ТС и тип сопоставленного пункта. Конфликт не угадывать, а записать CONFLICT."],
  [6,"Сопоставить пункты и пломбы","Использовать point_alias и точный seal_number. Несопоставленные строки оставить в фактах с UNKNOWN или NOT_FOUND."],
  [7,"Собрать факты","Создать все таблицы из output_tables без строк Итого и без объединённых ячеек."],
  [8,"Проверить","Выполнить quality_rules; AUTO плюс RAIL плюс UNKNOWN должно дать общий итог."],
  [9,"Сохранить","Сформировать один power_bi_dataset_<timestamp>.xlsx со стабильными именами таблиц."],
  [10,"Передать","Передать файл, журнал качества и версии трёх входных источников."]
], "AiHandoffTable", [8,30,104], "#2F75B5");

addSheet("input_files", "Входные файлы", "Три логических источника. Status_Automation используется только для подготовки третьего источника.", ["source_id","Русское название","Формат","Обязательность","Назначение"], [
  ["MAIN_EXPORT","Большая общая выгрузка перевозок","XLSX","Да","Единственный источник перевозок, включая авто и ЖД"],
  ["POINT_REGISTRY","Реестр пунктов с координатами","XLSX или JSON","Да","Канонические названия, варианты, транспорт, граница и координаты"],
  ["SEAL_SNAPSHOT","Единый снимок пломб","XLSX","Да для карты и пломб","Одна строка на пломбу, владелец, состояние, место, адрес и координаты"],
  ["STATUS_AUTOMATION","Проект подготовки снимка пломб","Папка проекта","Опционально","Формирует SEAL_SNAPSHOT из обработанных отчётов; не подключается Power BI напрямую"]
], "InputFilesTable", [22,38,22,26,80], "#5B9BD5");

const mapping = [
  ["MAIN_EXPORT","Дата создания","created_at","datetime","Да","Дата начала и фильтр второго этапа"], ["MAIN_EXPORT","Владелец","owner_name_raw","text","Нет","Исходный владелец"],
  ["MAIN_EXPORT","Номер перевозки","transportation_id","text","Да","Ключ перевозки"], ["MAIN_EXPORT","Статус","status_raw","text","Да","Исходный статус"],
  ["MAIN_EXPORT","Дата обновления","status_changed_at","datetime","Да","Дата появления текущего статуса"], ["MAIN_EXPORT","Тип перевозки","movement_type_raw","text","Да","Транзит, экспорт, взаимная торговля"],
  ["MAIN_EXPORT","Тип транспорта","transport_mode_raw","text","Да","Заявленный вид транспорта"], ["MAIN_EXPORT","Компания","company_name","text","Нет","Компания"], ["MAIN_EXPORT","БИН","company_bin","text","Нет","БИН текстом"],
  ["MAIN_EXPORT","Номера документов","document_numbers_raw","text","Нет","Исходные номера"], ["MAIN_EXPORT","Номера ТС","vehicle_numbers_raw","text","Нет","Номера через / проверяются по фрагментам"],
  ["MAIN_EXPORT","Коды ТН ВЭД","hs_codes_raw","text","Нет","Коды товаров"], ["MAIN_EXPORT","Номера НП","seal_numbers_raw","text","Нет","Номера пломб"],
  ["MAIN_EXPORT","Кол. путевых точек","route_point_count","whole number","Нет","Контроль маршрута"], ["MAIN_EXPORT","Кол. водителей","driver_count","whole number","Нет","Количество водителей"],
  ["MAIN_EXPORT","Кол. операторов","operator_count","whole number","Нет","Количество операторов"], ["MAIN_EXPORT","Начальный пункт","origin_point_raw","text","Да","Сопоставление с point_alias"],
  ["MAIN_EXPORT","Конечный пункт","destination_point_raw","text","Да","Сопоставление с point_alias"], ["MAIN_EXPORT","Страна начала","origin_country_raw","text","Да","ISO alpha 2"],
  ["MAIN_EXPORT","Страна завершения","destination_country_raw","text","Да","ISO alpha 2"], ["MAIN_EXPORT","Промежуточные точки","intermediate_points_raw","text","Нет","Маршрут и пункт наложения"],
  ["MAIN_EXPORT","Промежуточные страны","intermediate_countries_raw","text","Нет","Маршрут по странам"],
  ["POINT_REGISTRY","point_id","point_id","text","Да","Неизменный ключ"], ["POINT_REGISTRY","point_name_ru","point_name_ru","text","Да","Каноническое название"],
  ["POINT_REGISTRY","transport_mode_code","transport_mode_code","text","Да","AUTO или RAIL"], ["POINT_REGISTRY","border_type_code","border_type_code","text","Да","Внутренняя или внешняя граница ЕАЭС"],
  ["POINT_REGISTRY","latitude longitude","latitude longitude","decimal","Да","Координаты WGS84"], ["POINT_REGISTRY","aliases","alias_raw","text","Да","Варианты написания"],
  ["SEAL_SNAPSHOT","seal_number","seal_number","text","Да","Серийный номер пломбы"], ["SEAL_SNAPSHOT","owner_name","owner_name","text","Нет","Определённый владелец"],
  ["SEAL_SNAPSHOT","current_state_code","current_state_code","text","Да","Текущее состояние"], ["SEAL_SNAPSHOT","current_location_name","current_location_name","text","Нет","Пункт, склад, адрес или поселение"],
  ["SEAL_SNAPSHOT","latitude longitude","latitude longitude","decimal","Нет","Последние валидные координаты"], ["SEAL_SNAPSHOT","latest_fix_at","latest_fix_at","datetime","Нет","Дата последней координаты"]
];
addSheet("input_mapping", "Сопоставление входных столбцов", "Фактические заголовки приводятся к каноническим полям до расчёта показателей.", ["source_id","Исходный столбец","canonical_field","data_type","Обязательность","Описание"], mapping, "InputMappingTable", [20,34,32,18,20,70]);

const tables = [
  ["fact_transportation","Перевозки","Одна перевозка на дату среза","Да","Показатели перевозок и направления"], ["bridge_transportation_seal","Связь перевозок и пломб","Одна пломба в перевозке","Да","Связь с владельцем и состоянием пломбы"],
  ["fact_seal_snapshot","Снимок пломб","Одна пломба на дату среза","Да","Карта пломб, текущий адрес и состояние"], ["dim_point","Пункты","Один пункт одного вида транспорта","Да","Карта пунктов и фильтры"],
  ["point_alias","Варианты пунктов","Один вариант имени","Да","Нормализация названий"], ["dim_country","Страны","Одна страна","Да","Страны и направления"],
  ["dim_company","Компании","Одна компания","Да","Аналитика компаний"], ["dim_date","Календарь","Одна дата","Да","Временные фильтры"],
  ["dim_source","Источники","Одна версия входа","Да","Происхождение данных"], ["dictionary","Словарь кодов","Один код перечисления","Да","Русские подписи"],
  ["dq_issues","Проблемы качества","Одна проблема","Да","Диагностика"], ["metadata","Метаданные","Один параметр","Да","Дата и версия запуска"]
];
addSheet("output_tables", "Выходные таблицы", "Все листы результата создаются как именованные Excel Tables. Строки Итого в факты не добавляются.", ["table_name","Русское название","Детализация","Обязательность","Назначение"], tables, "OutputTablesTable", [32,34,58,22,65], "#70AD47");

const groups = {
  fact_transportation:["snapshot_date","transportation_id","created_at","status_changed_at","completed_at","status_code","status_group_code","movement_type_code","transport_mode_code","transport_mode_quality_code","origin_country_code","destination_country_code","counterparty_country_code","origin_point_id","destination_point_id","application_point_id","company_id","company_name","company_bin","route_point_count","driver_count","operator_count","source_id","quality_code"],
  bridge_transportation_seal:["snapshot_date","transportation_id","seal_number","seal_position","owner_country_code","is_kz_operator","match_status_code","source_id"],
  fact_seal_snapshot:["snapshot_at","seal_number","owner_name","owner_country_code","is_kz_operator","owner_quality_code","current_state_code","current_location_name","current_location_type_code","point_id","warehouse_id","warehouse_name","current_transportation_id","latest_fix_at","latitude","longitude","coordinate_source","nearest_country","nearest_settlement","location_quality_code","source_id"],
  dim_point:["point_id","site_id","point_name_ru","short_name_ru","transport_mode_code","border_type_code","neighbor_country_code","region_name_ru","district_name_ru","latitude","longitude","coordinate_source","active_flag"],
  point_alias:["alias_raw","source_id","point_id","match_status_code"], dim_country:["country_code","country_name_ru","eaeu_flag","sort_order"], dim_company:["company_id","company_name","company_bin"], dim_date:["date","year","month_number","month_name_ru","quarter"], dim_source:["source_id","source_type_code","file_name","file_modified_at","loaded_at","file_hash"], dictionary:["dictionary_name","code","display_name_ru","sort_order"], dq_issues:["issue_id","source_id","source_row_number","entity_type","entity_key","field_name","raw_value","issue_code","severity_code","comment"], metadata:["parameter_name","parameter_value"]
};
const descriptions = {snapshot_date:"Дата среза",snapshot_at:"Дата снимка",transportation_id:"Идентификатор перевозки",seal_number:"Серийный номер пломбы",created_at:"Дата создания перевозки",status_changed_at:"Дата появления текущего статуса",completed_at:"Дата завершения только для COMPLETED",latitude:"Широта WGS84",longitude:"Долгота WGS84",source_id:"Источник строки",quality_code:"Качество данных",point_id:"Ключ пункта",transport_mode_code:"AUTO RAIL или UNKNOWN"};
// Do not use a bare /count/ here: it also matches the beginning of
// `country_code`. Numeric names are matched only as complete tokens.
const typeFor = (f) => /(_at|_date|^date$)/.test(f) ? "datetime" : /(latitude|longitude)/.test(f) ? "decimal" : /(^|_)(count|number|position|year|month|quarter)(_|$)|(^|_)(sort_order|row_number)$/.test(f) ? "whole number" : /(flag|is_kz_operator)/.test(f) ? "boolean" : "text";
const dictRows = []; for (const [table, fields] of Object.entries(groups)) fields.forEach((f,i)=>dictRows.push([table,i+1,f,typeFor(f),["transportation_id","seal_number","snapshot_date","snapshot_at","source_id"].includes(f)?"Да":"Нет",descriptions[f]||"Поле согласно технической спецификации",f==="transportation_id"?"TR-000001":f==="seal_number"?"000002400006":f==="transport_mode_code"?"AUTO":""]));
addSheet("data_dictionary", "Словарь полей выходного файла", "Одна строка описывает одно физическое поле итогового Excel файла.", ["table_name","field_order","field_name","data_type","required","Описание","Пример"], dictRows, "DataDictionaryTable", [30,13,34,18,15,64,26]);

addSheet("metrics", "Показатели Power BI", "Меры создаются в Power BI на базе фактов, а не как строки Итого в Excel.", ["metric_code","Русское название","Таблица","Фильтр"], [
  ["TRANSPORTATIONS","Количество перевозок","fact_transportation","DISTINCTCOUNT transportation_id"], ["TRANSIT","Транзит","fact_transportation","movement_type_code = TRANSIT"], ["EXPORT","Экспорт из РК","fact_transportation","movement_type_code = EXPORT"],
  ["MUTUAL_TRADE","Взаимная торговля из РК","fact_transportation","movement_type_code = MUTUAL_TRADE"], ["COMPLETED_FROM_KZ","Завершено из РК","fact_transportation","COMPLETED; начало KZ; завершение не KZ"],
  ["COMPLETED_TO_KZ","Завершено в РК","fact_transportation","COMPLETED; начало не KZ; завершение KZ"], ["ACTIVE_TO_KZ","В сторону РК","fact_transportation","ACTIVE; завершение KZ"],
  ["SEALS_TOTAL","Пломбы в снимке","fact_seal_snapshot","DISTINCTCOUNT seal_number"], ["SEALS_IN_TRANSIT","Пломбы в пути","fact_seal_snapshot","current_state_code = IN_TRANSIT"], ["SEALS_IN_STOCK","Пломбы на складе","fact_seal_snapshot","current_state_code = IN_STOCK"], ["SEALS_OURS","Пломбы операторов РК","fact_seal_snapshot","is_kz_operator = TRUE"]
], "MetricsTable", [30,38,32,75], "#7030A0");

addSheet("code_lists", "Коды перечислений", "В фактах хранятся постоянные коды. Русские подписи задаются через этот словарь.", ["dictionary_name","code","display_name_ru","sort_order","Комментарий"], [
  ["transport_mode","AUTO","Автомобильный транспорт",1,""],["transport_mode","RAIL","Железнодорожный транспорт",2,""],["transport_mode","UNKNOWN","Не определено",99,"Не подменять авто"],
  ["movement_type","TRANSIT","Транзит",1,""],["movement_type","EXPORT","Экспорт из РК",2,""],["movement_type","MUTUAL_TRADE","Взаимная торговля из РК",3,""],["movement_type","OTHER","Прочая перевозка",98,""],["movement_type","UNKNOWN","Не определено",99,""],
  ["status_group","ACTIVE","Активная перевозка",1,""],["status_group","COMPLETED","Завершённая перевозка",2,""],["status_group","OTHER","Прочий статус",98,""],["status_group","UNKNOWN","Не определено",99,""],
  ["border_type","INTERNAL_EAEU","Внутренняя граница ЕАЭС",1,""],["border_type","EXTERNAL_EAEU","Внешняя граница ЕАЭС",2,""],
  ["seal_state","IN_TRANSIT","В перевозке",1,""],["seal_state","IN_STOCK","На складе",2,""],["seal_state","ISSUED","Выдана",3,""],["seal_state","OFFLINE","Нет связи",4,""],["seal_state","UNKNOWN","Не определено",99,""],
  ["quality","OK","Проверено",1,""],["quality","CONFLICT","Конфликт",2,"Требует проверки"],["quality","UNKNOWN","Не определено",99,""],["match_status","MATCHED","Сопоставлено",1,""],["match_status","NOT_FOUND","Не найдено",2,""],["match_status","AMBIGUOUS","Неоднозначно",3,""]
], "CodeListsTable", [28,28,40,14,62]);

addSheet("relationships", "Связи Power BI", "Связи однонаправленные от справочника к факту. Ролевые связи пунктов сделать неактивными.", ["Таблица один","Ключ","Таблица многие","Ключ","Кардинальность","Статус"], [
  ["dim_date","date","fact_transportation","snapshot_date","1:*","Активная"],["dim_date","date","fact_seal_snapshot","snapshot_at","1:*","Активная"],["dim_point","point_id","fact_transportation","application_point_id","1:*","Активная"],["dim_point","point_id","fact_seal_snapshot","point_id","1:*","Активная"],["dim_country","country_code","fact_transportation","counterparty_country_code","1:*","Активная"],["dim_company","company_id","fact_transportation","company_id","1:*","Активная"],["fact_transportation","transportation_id","bridge_transportation_seal","transportation_id","1:*","Активная"],["fact_seal_snapshot","seal_number","bridge_transportation_seal","seal_number","1:*","Активная"],["dim_source","source_id","Все факты","source_id","1:*","Активная"],["dim_point","point_id","point_alias","point_id","1:*","Служебная"]
], "RelationshipsTable", [28,30,34,32,18,28], "#A5A5A5");

const rawPoints = JSON.parse(await fs.readFile(pointsPath, "utf8")).points;
const pointRows = rawPoints.map(p=>{ const aliases=p.source_aliases||{}; return [p.id,p.id.replace(/_(auto|rail)$/, ""),p.name||p.report_name,p.report_name||p.name,p.transport==="авто"?"AUTO":"RAIL",p.border==="внутренняя"?"INTERNAL_EAEU":"EXTERNAL_EAEU","","","","",(p.aliases||[]).join(" | "),(aliases.big_export||[]).join(" | "),"Требуется заполнить координаты и проверить"]; });
const registry=addSheet("point_registry", "Реестр пунктов и станций", "Пункты перенесены из текущего JSON. Перед запуском необходимо заполнить координаты, страну и регион.", ["point_id","site_id","point_name_ru","short_name_ru","transport_mode_code","border_type_code","neighbor_country_code","region_name_ru","latitude","longitude","aliases","main_export_aliases","review_status"], pointRows, "PointRegistryTable", [27,24,28,28,25,27,24,28,16,16,55,55,55], "#4472C4");
registry.sh.getRange(`I6:J${registry.end}`).format.fill = AMBER; registry.sh.getRange(`I6:J${registry.end}`).format.numberFormat = "0.00000";

addSheet("quality_rules", "Проверки качества", "Проверки выполняются перед экспортом. Ошибка блокирует только зависимый факт или повреждённый файл.", ["rule_code","severity_code","Объект","Условие","Действие","blocking"], [
  ["DQ_STAGE_DATE","ERROR","fact_transportation","created_at >= 2026-07-31 00:00 UTC+05:00","Исключить из второго этапа и записать проблему","Да"],["DQ_DUP_TRANSPORT","ERROR","fact_transportation","snapshot_date + transportation_id уникальны","Оставить последнюю версию статуса","Да"],["DQ_TRANSPORT_TOTAL","ERROR","fact_transportation","AUTO + RAIL + UNKNOWN = общий итог","Показать расхождение","Да"],["DQ_POINT_FK","ERROR","Факты","Каждый point_id есть в dim_point","Записать исходное значение и UNKNOWN","Да"],["DQ_POINT_COORD","WARNING","dim_point","Активный пункт имеет валидные координаты","Не выводить на карту","Нет"],["DQ_COORD_RANGE","ERROR","fact_seal_snapshot","Широта -90..90, долгота -180..180, не 0,0","Очистить координаты и записать проблему","Да"],["DQ_SEAL_DUP","ERROR","fact_seal_snapshot","snapshot_at + seal_number уникальны","Оставить последнюю достоверную строку","Да"],["DQ_SEAL_MATCH","WARNING","bridge_transportation_seal","Пломба сопоставлена или NOT_FOUND","Не удалять связь","Нет"],["DQ_UNKNOWN_CODE","WARNING","Все факты","Коды есть в dictionary","Использовать UNKNOWN","Нет"],["DQ_TABLE_SCHEMA","ERROR","Выходной файл","Все обязательные таблицы и поля существуют","Прервать экспорт","Да"],["DQ_VALUE_TYPE","ERROR","Выходной файл","Даты и числа типизированы, идентификаторы текстовые","Исправить до экспорта","Да"]
], "QualityRulesTable", [28,18,32,80,65,16], "#A5A5A5");

addSheet("open_questions", "Данные для заполнения", "Поля отмечены как обязательные для полной карты и контролируются перед рабочим запуском.", ["question_id","Тема","Необходимые данные","Влияние","Статус"], [
  ["Q001","Координаты пунктов","Широта, долгота, регион, район, соседняя страна и источник координат для всех пунктов","Карта пунктов","Открыт"],["Q002","Снимок пломб","Единый SEAL_SNAPSHOT с владельцем, текущим местом, статусом и координатами","Карта и карточки пломб","Открыт"],["Q003","Складская принадлежность","Правило текущего склада и связь склад пункт пропуска при наличии","Пломбы на складах","Открыт"],["Q004","Словарь статусов","Полное соответствие исходного статуса ACTIVE COMPLETED OTHER","Показатели завершения","Открыт"],["Q005","Пороги свежести","Допустимый возраст снимка пломб в часах","Предупреждение STALE","Открыт"]
], "OpenQuestionsTable", [16,30,85,40,18], "#FFC000");

addSheet("status_automation", "Подготовка снимка пломб", "Используйте обработанные выходы Status_Automation. Прямое чтение RAW файлов в генератор Power BI не рекомендуется.", ["Приоритет","Файл или лист","Что берём","Результат в SEAL_SNAPSHOT"], [
  [1,"seal_latest_location_report.xlsx / Последние координаты","Серийный номер, страна принадлежности, дата, широта, долгота, источник, ближайшая страна и населённый пункт","Координаты и адрес пломбы"],
  [2,"seal_registry_all_operators_report.xlsx / Реестр","Определённый владелец, оператор, принадлежность оператору РК, качество и источник владельца","Владелец и качество"],
  [3,"traceability_report.xlsx / Текущее состояние","Текущее состояние, место, тип места, текущая перевозка, статус качества","Текущий статус, склад или пункт"],
  [4,"raw_01 и raw_02 только через адаптер","Складской срез и движения, если требуется уточнить склад","warehouse_id и warehouse_name"]
], "StatusAutomationTable", [16,52,72,62], "#5B9BD5");

for (const name of ["README","ai_handoff","input_files","input_mapping","output_tables","data_dictionary","metrics","code_lists","relationships","point_registry","quality_rules","open_questions","status_automation"]) wb.worksheets.getItem(name).getRange("A1:Z300").format.verticalAlignment = "center";
wb.recalculate(); await fs.mkdir(path.dirname(outputPath),{recursive:true}); await fs.mkdir(previewDir,{recursive:true});
const inspect = await wb.inspect({kind:"workbook,sheet,table",maxChars:12000,tableMaxRows:3,tableMaxCols:10,tableMaxCellChars:90}); console.log(inspect.ndjson);
const errors = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"final formula error scan"}); console.log(errors.ndjson);
for (const name of ["README","ai_handoff","input_files","input_mapping","output_tables","data_dictionary","metrics","code_lists","relationships","point_registry","quality_rules","open_questions","status_automation"]) { const sh=wb.worksheets.getItem(name); const used=sh.getUsedRange(); const blob=await wb.render({sheetName:name,range:`A1:${colName(Math.min(used.columnCount||8,13)-1)}${Math.min(used.rowCount||25,25)}`,scale:1,format:"png"}); await fs.writeFile(path.join(previewDir,`${name}.png`),new Uint8Array(await blob.arrayBuffer())); }
const out = await SpreadsheetFile.exportXlsx(wb); await out.save(outputPath); console.log(outputPath);
