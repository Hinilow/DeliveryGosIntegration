# Результаты SQL-запросов

База `sql/delivery.db` (SQLite), срез на 18.11.2026. Сами запросы с комментариями: [`sql/queries.sql`](queries.sql).

## 1. Готовность по командам

```sql
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
```

| команда | сервисов | в_проде | готовность_проц | прирост_за_неделю_пп |
|---|---|---|---|---|
| Платформа | 4 | 4 | 95.0 | 0.0 |
| Уведомления и статусы | 2 | 1 | 78.0 | 12.5 |
| Идентификация | 3 | 1 | 65.0 | 25.0 |
| Подпись и документы | 4 | 1 | 53.0 | 10.0 |
| СМЭВ-интеграции | 3 | 0 | 32.0 | 8.3 |
| Биометрия | 4 | 0 | 30.0 | 6.3 |

## 2. Динамика программы по неделям

```sql
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
```

| неделя | срез | в_проде | в_проде_за_неделю | готовность_проц | прирост_пп |
|---|---|---|---|---|---|
| W43 | 2026-10-21 | 3 |  | 28.5 |  |
| W44 | 2026-10-28 | 4 | 1 | 40.8 | 12.3 |
| W45 | 2026-11-04 | 4 | 0 | 47.0 | 6.2 |
| W46 | 2026-11-11 | 5 | 1 | 48.3 | 1.3 |
| W47 | 2026-11-18 | 7 | 2 | 57.8 | 9.5 |

## 3. Сколько недель сервис стоит на текущем этапе

```sql
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
```

| сервис | этап | недель_на_этапе | блокер |
|---|---|---|---|
| bio-template-store | Не начата | 5 | Нужна схема миграции stateful-хранилища |
| ebs-gateway | Анализ | 4 | Не согласованы требования ИБ к обработке биометрии в контейнерах |
| smev-adapter | Контейнеризация | 4 | Нет сертификата СМЭВ для контура DropApp |
| gosuslugi-status-push | Тест в DropApp | 3 | Ожидает сетевой доступ к ЕПГУ (заявка ИБ) |
| liveness-check | Контейнеризация | 3 | Нужны GPU-узлы в DropApp |
| request-status | Анализ | 3 | блокер не назван |
| sign-orchestrator | Тест в DropApp | 3 | блокер не назван |

## 4. Конфликты плана: сервис уходит в DropApp не позже своей зависимости

```sql
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
```

| сервис | план_сервиса | зависит_от | план_зависимости | этап_зависимости |
|---|---|---|---|---|
| sign-orchestrator | 2026-11-20 | doc-archive | 2026-12-18 | Анализ |
| sign-orchestrator | 2026-11-20 | ukep-adapter | 2026-12-04 | Контейнеризация |
| gosuslugi-status-push | 2026-11-20 | request-status | 2026-12-11 | Анализ |
| ebs-gateway | 2026-11-27 | bio-template-store | 2026-12-18 | Не начата |

## 5. Блокеры: сколько дней висят

```sql
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
```

| сервис | блокер | с_даты | календарных_дней | план |
|---|---|---|---|---|
| ebs-gateway | Не согласованы требования ИБ к обработке биометрии в контейнерах | 2026-10-28 | 21 | 2026-11-27 |
| smev-adapter | Нет сертификата СМЭВ для контура DropApp | 2026-10-30 | 19 | 2026-11-13 |
| gosuslugi-status-push | Ожидает сетевой доступ к ЕПГУ (заявка ИБ) | 2026-11-02 | 16 | 2026-11-20 |
| liveness-check | Нужны GPU-узлы в DropApp | 2026-11-03 | 15 | 2026-12-11 |
| bio-template-store | Нужна схема миграции stateful-хранилища | 2026-11-05 | 13 | 2026-12-18 |
| ukep-adapter | Нужен доступ к HSM из контура DropApp | 2026-11-12 | 6 | 2026-12-04 |

## 6. В проде, но с долгами: открытые пункты чек-листа или РЦОД не готов

```sql
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
```

| сервис | долги |
|---|---|
| consent-registry | РЦОД не готов |
| notification-hub | приёмка сопровождением |
| observability-agent | runbook и план отката, РЦОД не готов |

## 7. Старый контур после выхода в прод

```sql
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
```

| сервис | план | вышел_в_прод | отклонение_дн | недель_в_двух_контурах |
|---|---|---|---|---|
| audit-log | 2026-10-23 | 2026-10-16 | -7 | 4 |
| observability-agent | 2026-10-23 | 2026-10-23 | 0 | 3 |
| doc-render | 2026-11-06 | 2026-11-06 | 0 | 1 |
| consent-registry | 2026-11-13 | 2026-11-13 | 0 | 0 |
| notification-hub | 2026-11-06 | 2026-11-13 | 7 | 0 |

## 8. Готовность РЦОД по командам

```sql
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
```

| команда | сервисов | развёрнуто_в_рцод | учения_пройдены | учения_не_пройдены |
|---|---|---|---|---|
| Платформа | 4 | 3 | 3 | 0 |
| Подпись и документы | 4 | 1 | 1 | 0 |
| Уведомления и статусы | 2 | 1 | 1 | 0 |
| Идентификация | 3 | 1 | 0 | 1 |
| Биометрия | 4 | 0 | 0 |  |
| СМЭВ-интеграции | 3 | 0 | 0 |  |

## 9. Сервисы с внешними ГИС: этап миграции и РЦОД

```sql
SELECT s.service AS сервис, s.external_system AS гис, n.stage_now AS этап, r.rcod AS рцод, s.peak_rps AS пик_rps
FROM services s
JOIN stage_now n ON n.service_id = s.service_id
JOIN rcod r      ON r.service_id = s.service_id
WHERE s.external_system IS NOT NULL
ORDER BY s.peak_rps DESC;
```

| сервис | гис | этап | рцод | пик_rps |
|---|---|---|---|---|
| esia-auth-adapter | ЕСИА | Тест в DropApp | Частично | 1500 |
| smev-adapter | СМЭВ | Контейнеризация | Нет | 1100 |
| ebs-gateway | ЕБС | Анализ | Нет | 400 |
| gosuslugi-status-push | ЕПГУ | Тест в DropApp | Частично | 350 |
| ukep-adapter | Госключ | Контейнеризация | Нет | 300 |
