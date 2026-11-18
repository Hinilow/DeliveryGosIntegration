-- Запросы к sql/delivery.db. Каждый блок начинается с собачки и заголовка, scripts/build.py выполняет
-- их по очереди и пишет результаты в sql/results.md.

-- @ 1. Готовность по командам
SELECT t.team                                             AS команда,
       COUNT(*)                                           AS сервисов,
       SUM(st.readiness >= 0.9)                           AS в_проде,
       ROUND(100.0 * AVG(st.readiness))                   AS готовность_проц,
       ROUND(100.0 * AVG(st.readiness - sp.readiness), 1) AS прирост_за_неделю_пп
FROM stage_now n
JOIN services s ON s.service_id = n.service_id
JOIN teams t    ON t.team_id    = s.team_id
JOIN stages st  ON st.stage     = n.stage_now
JOIN stages sp  ON sp.stage     = n.stage_prev
GROUP BY t.team
ORDER BY готовность_проц DESC;

-- @ 2. Динамика программы по неделям
-- Темп считаю по последним неделям, а не по всей истории: в начале переезжали простые сервисы Платформы.
WITH weekly AS (
    SELECT h.week, h.snapshot_date,
           SUM(st.readiness >= 0.9)        AS в_проде,
           ROUND(100.0 * AVG(st.readiness), 1) AS готовность_проц
    FROM history h
    JOIN stages st ON st.stage = h.stage
    GROUP BY h.week, h.snapshot_date
)
SELECT week AS неделя, snapshot_date AS срез, в_проде,
       в_проде - LAG(в_проде) OVER (ORDER BY week)                   AS в_проде_за_неделю,
       готовность_проц,
       ROUND(готовность_проц - LAG(готовность_проц) OVER (ORDER BY week), 2) AS прирост_пп
FROM weekly
ORDER BY week;

-- @ 3. Сколько недель сервис стоит на текущем этапе
-- Классическая задача gaps and islands: разница двух ROW_NUMBER одинакова внутри непрерывного отрезка
-- на одном этапе. Сервис, который три недели не двигается, под риском, даже если команда не назвала блокер.
WITH h AS (
    SELECT service_id, week, stage,
           ROW_NUMBER() OVER (PARTITION BY service_id ORDER BY week)
         - ROW_NUMBER() OVER (PARTITION BY service_id, stage ORDER BY week) AS grp
    FROM history
),
runs AS (
    SELECT service_id, stage, grp, COUNT(*) AS weeks, MAX(week) AS last_week
    FROM h
    GROUP BY service_id, stage, grp
)
SELECT s.service                              AS сервис,
       r.stage                                AS этап,
       r.weeks                                AS недель_на_этапе,
       COALESCE(m.blocker, 'блокер не назван') AS блокер
FROM runs r
JOIN services s  ON s.service_id = r.service_id
JOIN migration m ON m.service_id = r.service_id
JOIN stages st   ON st.stage     = r.stage
WHERE r.last_week = (SELECT MAX(week) FROM history)
  AND st.readiness < 0.9
  AND r.weeks >= 3
ORDER BY r.weeks DESC, s.service;

-- @ 4. Конфликты плана: сервис уходит в DropApp не позже своей зависимости
SELECT s.service    AS сервис,
       ma.plan_prod AS план_сервиса,
       d2.service   AS зависит_от,
       mb.plan_prod AS план_зависимости,
       nb.stage_now AS этап_зависимости
FROM dependencies d
JOIN services s   ON s.service_id  = d.service_id
JOIN services d2  ON d2.service_id = d.depends_on
JOIN migration ma ON ma.service_id = d.service_id
JOIN migration mb ON mb.service_id = d.depends_on
JOIN stage_now nb ON nb.service_id = d.depends_on
JOIN stages sb    ON sb.stage      = nb.stage_now
WHERE sb.readiness < 0.9
  AND ma.plan_prod <= mb.plan_prod
ORDER BY ma.plan_prod;

-- @ 5. Блокеры: сколько дней висят
WITH params AS (SELECT DATE('2026-11-18') AS report_date)
SELECT s.service                                                         AS сервис,
       m.blocker                                                         AS блокер,
       m.blocker_since                                                   AS с_даты,
       CAST(JULIANDAY(p.report_date) - JULIANDAY(m.blocker_since) AS INTEGER) AS календарных_дней,
       m.plan_prod                                                       AS план
FROM migration m
JOIN services s ON s.service_id = m.service_id
CROSS JOIN params p
WHERE m.blocker IS NOT NULL
ORDER BY m.blocker_since;

-- @ 6. В проде, но с долгами: открытые пункты чек-листа или РЦОД не готов
SELECT s.service AS сервис,
       TRIM(COALESCE(GROUP_CONCAT(c.item, ', '), '') ||
            CASE WHEN r.rcod <> 'Да' OR r.dr_result IS NOT 'Успех' OR r.rto_fact_min > r.rto_target_min
                 THEN CASE WHEN COUNT(c.item) > 0 THEN ', ' ELSE '' END || 'РЦОД не готов' ELSE '' END) AS долги
FROM stage_now n
JOIN services s  ON s.service_id = n.service_id
JOIN stages st   ON st.stage     = n.stage_now
JOIN rcod r      ON r.service_id = n.service_id
LEFT JOIN checklist c ON c.service_id = n.service_id AND c.done = 0
WHERE st.readiness >= 0.9
GROUP BY s.service, r.rcod, r.dr_result, r.rto_fact_min, r.rto_target_min
HAVING долги <> ''
ORDER BY s.service;

-- @ 7. Старый контур после выхода в прод
-- Пока старый контур не выведен, платим за два контура. Две недели держим ради отката, дальше это хвост.
WITH params AS (SELECT DATE('2026-11-18') AS report_date)
SELECT s.service                                                             AS сервис,
       m.plan_prod                                                           AS план,
       m.prod_fact                                                           AS вышел_в_прод,
       CAST(JULIANDAY(m.prod_fact) - JULIANDAY(m.plan_prod) AS INTEGER)      AS отклонение_дн,
       CAST((JULIANDAY(p.report_date) - JULIANDAY(m.prod_fact)) / 7 AS INTEGER) AS недель_в_двух_контурах
FROM migration m
JOIN services s  ON s.service_id = m.service_id
JOIN stage_now n ON n.service_id = m.service_id
CROSS JOIN params p
WHERE n.stage_now = 'Прод в DropApp'
ORDER BY недель_в_двух_контурах DESC;

-- @ 8. Готовность РЦОД по командам
SELECT t.team AS команда,
       COUNT(*) AS сервисов,
       SUM(r.rcod = 'Да')                                                                   AS развёрнуто_в_рцод,
       SUM(r.rcod = 'Да' AND r.dr_result = 'Успех' AND r.rto_fact_min <= r.rto_target_min) AS учения_пройдены,
       SUM(r.dr_result = 'Неуспех')                                                         AS учения_не_пройдены
FROM rcod r
JOIN services s ON s.service_id = r.service_id
JOIN teams t    ON t.team_id    = s.team_id
GROUP BY t.team
ORDER BY учения_пройдены DESC, развёрнуто_в_рцод DESC;

-- @ 9. Сервисы с внешними ГИС: этап миграции и РЦОД
SELECT s.service AS сервис, s.external_system AS гис, n.stage_now AS этап, r.rcod AS рцод, s.peak_rps AS пик_rps
FROM services s
JOIN stage_now n ON n.service_id = s.service_id
JOIN rcod r      ON r.service_id = s.service_id
WHERE s.external_system IS NOT NULL
ORDER BY s.peak_rps DESC;
