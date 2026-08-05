# modiin_weeks.csv — הסבר עמודות / Column Guide / Guia de Colunas

> **קובץ זה הוא מקור הזמנים לכל השנה** — כל לוח חודשי נבנה ממנו.
> **This file is the source of all prayer times for the year** — every monthly calendar is built from it.
> **Este arquivo é a fonte de todos os horários do ano** — cada calendário mensal é gerado a partir dele.

---

## עברית

**מקור הזמנים:**
- **הדלקת נרות (candles), שקיעה (sunset_rabanut), צאת השבת (motzash)** — נלקחים ישירות מלוח הזמנים הרשמי של **הרבנות מודיעין**. אינם מחושבים ולעולם לא משתנים על ידי הכלי.
- שאר זמני השמש (שקיעות אמצע שבוע, סוף זמן ק"ש) מחושבים בתוכנה (ephem), מכוילים מול הרבנות.

**עמודות ה-delta** = הפרש לבדיקה בלבד (דקות שלמות). מודדים כמה החישוב שלנו סוטה מהרבנות. צריכים להיות קרובים ל-0; משמשים כבקרת איכות, לא מוצגים בלוח.

**calc מול adj**: לכל זמן מחושב יש עמודת `_calc` (מחושבת מחדש בכל הרצה) ועמודת `_adj` (עריכה ידנית). אם `_adj` מלא — הוא גובר. השאר ריק כדי להשתמש בחישוב.

| # | עמודה | תיאור |
|---|---|---|
| 1 | week | מספר שבוע רץ (1..N) |
| 2 | hyrmo | מפתח שנה+חודש (תשרי=01), למשל 578612 = אלול תשפ"ו |
| 3 | heb_year | שנה עברית כמספר, למשל 5786 |
| 4 | heb_year_str | שנה עברית באותיות, תשפ"ו |
| 5 | heb_month_num | מספר חודש, תשרי=1 (כולל אדר א'/ב' בשנה מעוברת) |
| 6 | heb_month_name | שם החודש, למשל אלול / אדר א' |
| 7 | from | יום ראשון של השבוע. כל זמני החול בשורה (שקיעה מוקדמת, מנחת חול) מחושבים עבור הימים שבין from ל-to |
| 8 | to | יום חמישי של אותו שבוע (סוף טווח ימי החול א'–ה') |
| 9 | weekend_name | שם הפרשה של אותה שבת |
| 10 | weekend_date | תאריך השבת (העוגן של השורה) |
| 11 | taarich | תאריך עברי של השבת, למשל ה' תשרי |
| 12 | **candles** | **הדלקת נרות — מהרבנות מודיעין** |
| 13 | **motzash** | **צאת השבת — מהרבנות מודיעין** |
| 14 | sharbat_mincha | yes/no — מנחת שרבט בקיץ 17:15 ("yes" אם הדלקת נרות אחרי 18:00) |
| 15 | shabbat_mincha_ktana_calc | מנחה קטנה בשבת = עיגול כלפי מטה ל-5 דק' של הדלקת נרות |
| 16 | shabbat_mincha_ktana_adj | עקיפה ידנית (גובר על calc) |
| 17 | shacharit1 | שחרית א' — ריק = ברירת מחדל 6:45; מלא = עוקף (למשל ר"ח 6:30) |
| 18 | sof_zman_shma_calc | סוף זמן ק"ש (גר"א) — מחושב ephem, מכויל |
| 19 | sof_zman_shma_adj | עקיפה ידנית (גובר על calc) |
| 20 | **sunset_rabanut** | **שקיעה רשמית של השבת — מהרבנות מודיעין** |
| 21 | sunset_ephem | שקיעה של השבת כפי שחישבה התוכנה |
| 22 | sunset_delta | פער בין השקיעה לפי זמני הרבנות לבין הזמנים המחושבים. מיועד לבקרת איכות של הזמנים המחושבים |
| 23 | earliest_sunset_weekday | השקיעה המוקדמת ביותר א'–ה' (בסיס למנחת חול) |
| 24 | weekday_sunset_delta | פער בין השקיעה המאוחרת למוקדמת בחול (דקות). נותן אפשרות לבקרה נוספת |
| 25 | weekday_mincha_calc | מנחה בחול = עיגול מטה ל-5 דק' של (השקיעה המוקדמת − 12 דק') |
| 26 | weekday_mincha_adj | אפשרות למעקף במקרה הצורך לזמן המחושב של מנחה. מיועד לשינויים חריגים בלבד מהחישוב |

**קבועים (לא בקובץ):** מנחה גדולה 13:30 · מנחת שרבט 17:15 · שחרית ב' 8:30 · מנחת ערב שבת = מנחה קטנה + 10 דק' (מחושב בזמן ההפקה).

---

## English

**Source of the times:**
- **הדלקת נרות (candles), שקיעה (sunset_rabanut), מוצאי שבת (motzash)** are taken directly from the official timetable of the **Modiin Rabbinate (הרבנות מודיעין)**. They are never calculated and never modified by the tool.
- The other sun times (weekday שקיעה, סוף זמן קריאת שמע) are computed in software (ephem), calibrated against the Rabbinate.

**"delta" columns** = a check only (whole minutes) for quality-control of the computed times — they are not printed on the calendar.

**calc vs adj**: each computed time has a `_calc` column (recomputed every run) and an `_adj` column (manual edit). If `_adj` is filled it wins; leave it blank to use the computed value.

| # | Column | Description |
|---|---|---|
| 1 | week | Sequential week number (1..N) |
| 2 | hyrmo | Year+month key (תשרי=01), e.g. 578612 = אלול 5786 |
| 3 | heb_year | Hebrew year as a number, e.g. 5786 |
| 4 | heb_year_str | Hebrew year in letters (תשפ"ו) |
| 5 | heb_month_num | Month number, תשרי=1 (handles leap-year אדר א'/ב') |
| 6 | heb_month_name | Month name, e.g. אלול / אדר א' |
| 7 | from | Sunday of the week. All weekday times in the row (earliest שקיעה, weekday מנחה) are computed for the days between from and to |
| 8 | to | Thursday of that week (end of the א'–ה' weekday range) |
| 9 | weekend_name | פרשה name for that שבת |
| 10 | weekend_date | The שבת date (the row's anchor) |
| 11 | taarich | Hebrew date of the שבת, e.g. ה' תשרי |
| 12 | **candles** | **הדלקת נרות — from the Modiin Rabbinate** |
| 13 | **motzash** | **מוצאי שבת — from the Modiin Rabbinate** |
| 14 | sharbat_mincha | yes/no — summer מנחת שרבט 17:15 ("yes" if הדלקת נרות is after 18:00) |
| 15 | shabbat_mincha_ktana_calc | מנחה קטנה on שבת = הדלקת נרות floored to 5 min |
| 16 | shabbat_mincha_ktana_adj | Manual override (wins over calc) |
| 17 | shacharit1 | שחרית א' — blank = default 6:45; fill to override (e.g. ראש חודש 6:30) |
| 18 | sof_zman_shma_calc | סוף זמן קריאת שמע (גר"א) — ephem-computed, calibrated |
| 19 | sof_zman_shma_adj | Manual override (wins over calc) |
| 20 | **sunset_rabanut** | **Official שבת שקיעה — from the Modiin Rabbinate** |
| 21 | sunset_ephem | שבת שקיעה as computed by the software |
| 22 | sunset_delta | Difference between the Rabbinate שקיעה and the computed time. Quality-control of the computed times |
| 23 | earliest_sunset_weekday | Earliest שקיעה א'–ה' (basis for weekday מנחה) |
| 24 | weekday_sunset_delta | Spread between the latest and earliest weekday שקיעה (minutes). Provides an additional check |
| 25 | weekday_mincha_calc | Weekday מנחה = (earliest שקיעה − 12 min) floored to 5 min |
| 26 | weekday_mincha_adj | Optional override of the computed מנחה time when needed. For exceptional deviations from the calculation only |

**Constants (not in the file):** מנחה גדולה 13:30 · מנחת שרבט 17:15 · שחרית ב' 8:30 · מנחת ערב שבת = מנחה קטנה + 10 min (computed at render time).

---

## Português (Brasil)

**Origem dos horários:**
- **הדלקת נרות (candles), שקיעה / pôr do sol (sunset_rabanut), מוצאי שבת (motzash)** são retirados diretamente da tabela oficial do **Rabinato de Modiin (הרבנות מודיעין)**. Nunca são calculados nem alterados pela ferramenta.
- Os demais horários solares (שקיעה nos dias de semana, סוף זמן קריאת שמע) são calculados por software (ephem), calibrados em relação ao Rabinato.

**Colunas "delta"** = apenas uma verificação (minutos inteiros) para controle de qualidade dos horários calculados — não aparecem no calendário.

**calc vs adj**: cada horário calculado tem uma coluna `_calc` (recalculada a cada execução) e uma coluna `_adj` (edição manual). Se `_adj` estiver preenchida, ela prevalece; deixe em branco para usar o valor calculado.

| # | Coluna | Descrição |
|---|---|---|
| 1 | week | Número sequencial da semana (1..N) |
| 2 | hyrmo | Chave ano+mês (תשרי=01), ex. 578612 = אלול 5786 |
| 3 | heb_year | Ano hebraico como número, ex. 5786 |
| 4 | heb_year_str | Ano hebraico em letras (תשפ"ו) |
| 5 | heb_month_num | Número do mês, תשרי=1 (trata אדר א'/ב' em ano bissexto) |
| 6 | heb_month_name | Nome do mês, ex. אלול / אדר א' |
| 7 | from | Domingo da semana. Todos os horários dos dias de semana na linha (שקיעה mais cedo, מנחה de dia de semana) são calculados para os dias entre from e to |
| 8 | to | Quinta-feira da mesma semana (fim do intervalo de dias א'–ה') |
| 9 | weekend_name | Nome da פרשה daquele שבת |
| 10 | weekend_date | Data do שבת (âncora da linha) |
| 11 | taarich | Data hebraica do שבת, ex. ה' תשרי |
| 12 | **candles** | **הדלקת נרות — do Rabinato de Modiin** |
| 13 | **motzash** | **מוצאי שבת — do Rabinato de Modiin** |
| 14 | sharbat_mincha | yes/no — מנחת שרבט de verão 17:15 ("yes" se הדלקת נרות for após 18:00) |
| 15 | shabbat_mincha_ktana_calc | מנחה קטנה no שבת = הדלקת נרות arredondado p/ baixo a 5 min |
| 16 | shabbat_mincha_ktana_adj | Ajuste manual (prevalece sobre calc) |
| 17 | shacharit1 | שחרית א' — vazio = padrão 6:45; preencher para substituir (ex. ראש חודש 6:30) |
| 18 | sof_zman_shma_calc | סוף זמן קריאת שמע (גר"א) — calculado por ephem, calibrado |
| 19 | sof_zman_shma_adj | Ajuste manual (prevalece sobre calc) |
| 20 | **sunset_rabanut** | **שקיעה oficial do שבת — do Rabinato de Modiin** |
| 21 | sunset_ephem | שקיעה do שבת calculado pelo software |
| 22 | sunset_delta | Diferença entre a שקיעה do Rabinato e o horário calculado. Controle de qualidade dos horários calculados |
| 23 | earliest_sunset_weekday | שקיעה mais cedo א'–ה' (base para a מנחה de dia de semana) |
| 24 | weekday_sunset_delta | Diferença entre a שקיעה mais tarde e a mais cedo nos dias de semana (minutos). Oferece uma verificação adicional |
| 25 | weekday_mincha_calc | מנחה de dia de semana = (שקיעה mais cedo − 12 min) arredondado p/ baixo a 5 min |
| 26 | weekday_mincha_adj | Substituição opcional do horário calculado da מנחה quando necessário. Apenas para desvios excepcionais do cálculo |

**Constantes (não estão no arquivo):** מנחה גדולה 13:30 · מנחת שרבט 17:15 · שחרית ב' 8:30 · מנחת ערב שבת = מנחה קטנה + 10 min (calculada na geração).
