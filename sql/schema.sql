-- База трекера миграции кластера госинтеграций на DropApp (SQLite 3.25+, нужны оконные функции).
-- Этапы хранятся как недельная история, текущий этап и этап неделю назад берутся из представления stage_now.

CREATE TABLE stages (
    stage       TEXT PRIMARY KEY,
    stage_order INTEGER NOT NULL,
    readiness   REAL    NOT NULL           -- доля готовности этапа, 0..1
);

CREATE TABLE teams (
    team_id   TEXT PRIMARY KEY,
    team      TEXT NOT NULL,
    lead_role TEXT NOT NULL,
    focus     TEXT
);

CREATE TABLE services (
    service_id      TEXT PRIMARY KEY,
    service         TEXT NOT NULL UNIQUE,
    team_id         TEXT NOT NULL REFERENCES teams(team_id),
    purpose         TEXT,
    criticality     TEXT,
    peak_rps        INTEGER,
    external_system TEXT                  -- внешняя ГИС, если сервис с ней работает
);

CREATE TABLE dependencies (
    service_id TEXT NOT NULL REFERENCES services(service_id),
    depends_on TEXT NOT NULL REFERENCES services(service_id),
    PRIMARY KEY (service_id, depends_on)
);

-- Срез этапов раз в неделю, по средам перед встречей руководителей групп
CREATE TABLE history (
    week          TEXT NOT NULL,          -- W43, W44, ...
    snapshot_date DATE NOT NULL,
    service_id    TEXT NOT NULL REFERENCES services(service_id),
    stage         TEXT NOT NULL REFERENCES stages(stage),
    PRIMARY KEY (week, service_id)
);

CREATE TABLE migration (
    service_id    TEXT PRIMARY KEY REFERENCES services(service_id),
    plan_prod     DATE NOT NULL,          -- план выхода в прод DropApp
    prod_fact     DATE,                   -- фактический выход в прод
    owner_role    TEXT NOT NULL,
    blocker       TEXT,
    blocker_since DATE,
    comment       TEXT
);

-- Чек-лист выхода в прод: одна строка на сервис и пункт
CREATE TABLE checklist (
    service_id TEXT NOT NULL REFERENCES services(service_id),
    item       TEXT NOT NULL,
    done       INTEGER NOT NULL CHECK (done IN (0, 1)),
    PRIMARY KEY (service_id, item)
);

CREATE TABLE rcod (
    service_id     TEXT PRIMARY KEY REFERENCES services(service_id),
    main_dc        TEXT NOT NULL,
    rcod           TEXT NOT NULL CHECK (rcod IN ('Да', 'Частично', 'Нет')),
    dr_test_date   DATE,
    dr_result      TEXT CHECK (dr_result IN ('Успех', 'Неуспех')),
    rto_target_min INTEGER NOT NULL,
    rto_fact_min   INTEGER
);

CREATE TABLE calendar (
    date_from DATE NOT NULL,
    date_to   DATE NOT NULL,
    event     TEXT NOT NULL,
    kind      TEXT NOT NULL
);

CREATE VIEW stage_now AS
WITH h AS (
    SELECT service_id, week, stage,
           LAG(stage) OVER (PARTITION BY service_id ORDER BY week) AS stage_prev
    FROM history
)
SELECT service_id, stage AS stage_now, stage_prev
FROM h
WHERE week = (SELECT MAX(week) FROM history);
