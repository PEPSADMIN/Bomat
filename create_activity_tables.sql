-- ============================================================
-- PEPS BOM Tool — Activity & History Tables
-- Database : BOM  (192.168.0.133)
-- Run once  : Execute in SSMS against the BOM database
-- ============================================================

USE BOM;
GO

-- ============================================================
-- 1. BOM_UserActivity
--    Every login, download, replace, approval event
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'BOM_UserActivity')
BEGIN
    CREATE TABLE BOM_UserActivity (
        id          INT          IDENTITY(1,1) PRIMARY KEY,
        username    NVARCHAR(100) NOT NULL,
        user_role   NVARCHAR(50)  NOT NULL DEFAULT '',
        activity    NVARCHAR(100) NOT NULL,
        -- e.g. login | logout | bom_download | global_replace
        --      approval_submit | approval_approve | approval_reject
        --      bom_edit | nearest_bom | new_product
        detail      NVARCHAR(1000) NOT NULL DEFAULT '',
        ip_address  NVARCHAR(50)   NOT NULL DEFAULT '',
        session_id  NVARCHAR(100)  NOT NULL DEFAULT '',
        created_at  DATETIME2     NOT NULL DEFAULT GETDATE()
    );
    CREATE INDEX IX_UserActivity_Username  ON BOM_UserActivity (username);
    CREATE INDEX IX_UserActivity_Activity  ON BOM_UserActivity (activity);
    CREATE INDEX IX_UserActivity_CreatedAt ON BOM_UserActivity (created_at DESC);
    PRINT 'Created BOM_UserActivity';
END
ELSE
    PRINT 'BOM_UserActivity already exists — skipped';
GO

-- ============================================================
-- 2. BOM_RunHistory
--    One row per BOM generation run
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'BOM_RunHistory')
BEGIN
    CREATE TABLE BOM_RunHistory (
        id              INT           IDENTITY(1,1) PRIMARY KEY,
        sqlite_run_id   INT           NOT NULL DEFAULT 0,   -- mirrors runs.id in local SQLite
        run_label       NVARCHAR(400) NOT NULL,
        status          NVARCHAR(20)  NOT NULL DEFAULT 'OK',
        product_count   INT           NOT NULL DEFAULT 0,
        rows_generated  INT           NOT NULL DEFAULT 0,
        run_by          NVARCHAR(100) NOT NULL DEFAULT '',
        mode            NVARCHAR(50)  NOT NULL DEFAULT 'FULL',
        output_mode     NVARCHAR(50)  NOT NULL DEFAULT 'ZIP',
        output_filename NVARCHAR(400) NOT NULL DEFAULT '',
        approval_ref    NVARCHAR(200) NOT NULL DEFAULT '',
        warnings        INT           NOT NULL DEFAULT 0,
        created_at      DATETIME2     NOT NULL DEFAULT GETDATE(),
        completed_at    DATETIME2     NULL
    );
    CREATE INDEX IX_RunHistory_RunBy     ON BOM_RunHistory (run_by);
    CREATE INDEX IX_RunHistory_CreatedAt ON BOM_RunHistory (created_at DESC);
    PRINT 'Created BOM_RunHistory';
END
ELSE
    PRINT 'BOM_RunHistory already exists — skipped';
GO

-- ============================================================
-- 3. BOM_GlobalReplace
--    One row per global material-code replacement executed
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'BOM_GlobalReplace')
BEGIN
    CREATE TABLE BOM_GlobalReplace (
        id              INT            IDENTITY(1,1) PRIMARY KEY,
        sqlite_snap_id  INT            NOT NULL DEFAULT 0,
        old_code        NVARCHAR(300)  NOT NULL,
        new_code        NVARCHAR(300)  NOT NULL,
        reason          NVARCHAR(1000) NOT NULL DEFAULT '',
        approval_ref    NVARCHAR(200)  NOT NULL DEFAULT '',
        run_by          NVARCHAR(100)  NOT NULL DEFAULT '',
        files_changed   INT            NOT NULL DEFAULT 0,
        rows_changed    INT            NOT NULL DEFAULT 0,
        status          NVARCHAR(50)   NOT NULL DEFAULT 'active',
        created_at      DATETIME2      NOT NULL DEFAULT GETDATE()
    );
    CREATE INDEX IX_GlobalReplace_OldCode   ON BOM_GlobalReplace (old_code);
    CREATE INDEX IX_GlobalReplace_RunBy     ON BOM_GlobalReplace (run_by);
    CREATE INDEX IX_GlobalReplace_CreatedAt ON BOM_GlobalReplace (created_at DESC);
    PRINT 'Created BOM_GlobalReplace';
END
ELSE
    PRINT 'BOM_GlobalReplace already exists — skipped';
GO

-- ============================================================
-- 4. BOM_ApprovalHistory
--    All approval requests and their outcomes
-- ============================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'BOM_ApprovalHistory')
BEGIN
    CREATE TABLE BOM_ApprovalHistory (
        id                 INT            IDENTITY(1,1) PRIMARY KEY,
        sqlite_approval_id INT            NOT NULL DEFAULT 0,
        username           NVARCHAR(100)  NOT NULL,
        user_role          NVARCHAR(50)   NOT NULL DEFAULT '',
        action_type        NVARCHAR(100)  NOT NULL,
        action_summary     NVARCHAR(1000) NOT NULL DEFAULT '',
        status             NVARCHAR(50)   NOT NULL DEFAULT 'pending',
        reviewer_name      NVARCHAR(100)  NOT NULL DEFAULT '',
        rejection_reason   NVARCHAR(1000) NOT NULL DEFAULT '',
        created_at         DATETIME2      NOT NULL DEFAULT GETDATE(),
        reviewed_at        DATETIME2      NULL
    );
    CREATE INDEX IX_ApprovalHistory_Username  ON BOM_ApprovalHistory (username);
    CREATE INDEX IX_ApprovalHistory_Status    ON BOM_ApprovalHistory (status);
    CREATE INDEX IX_ApprovalHistory_CreatedAt ON BOM_ApprovalHistory (created_at DESC);
    PRINT 'Created BOM_ApprovalHistory';
END
ELSE
    PRINT 'BOM_ApprovalHistory already exists — skipped';
GO

-- ============================================================
-- 5. VW_UserSessionHistory
--    Session timeline: each login paired with its events
--    up to the next logout.  Query this view in SSMS to see
--    everything a user did between 8:30 login and 6:30 logout.
--
--    Example:
--      SELECT * FROM VW_UserSessionHistory
--      WHERE username = 'Hari'
--        AND CAST(login_at AS DATE) = '2026-06-13'
--      ORDER BY login_at, event_seq;
-- ============================================================
IF EXISTS (SELECT 1 FROM sys.views WHERE name = 'VW_UserSessionHistory')
    DROP VIEW VW_UserSessionHistory;
GO

CREATE VIEW VW_UserSessionHistory AS
WITH numbered AS (
    -- Assign a sequential row number per user ordered by time
    SELECT
        id,
        username,
        user_role,
        activity,
        detail,
        ip_address,
        session_id,
        created_at,
        ROW_NUMBER() OVER (PARTITION BY username ORDER BY created_at, id) AS rn
    FROM BOM_UserActivity
),
logins AS (
    -- Every login row becomes the start of a session
    SELECT
        id            AS login_id,
        username,
        user_role,
        ip_address,
        session_id,
        created_at    AS login_at,
        rn            AS login_rn,
        -- Find the rn of the NEXT login for this user (used to bound the session)
        LEAD(rn) OVER (PARTITION BY username ORDER BY rn) AS next_login_rn
    FROM numbered
    WHERE activity = 'login'
)
SELECT
    l.username,
    l.user_role,
    l.login_at,
    (
        SELECT TOP 1 created_at
        FROM numbered n2
        WHERE n2.username = l.username
          AND n2.activity = 'logout'
          AND n2.rn > l.login_rn
          AND (l.next_login_rn IS NULL OR n2.rn < l.next_login_rn)
        ORDER BY n2.rn
    )                          AS logout_at,
    l.ip_address,
    n.activity,
    n.detail,
    n.created_at               AS event_at,
    n.rn - l.login_rn          AS event_seq
FROM logins l
JOIN numbered n
  ON  n.username = l.username
  AND n.rn > l.login_rn
  AND (l.next_login_rn IS NULL OR n.rn < l.next_login_rn)
WHERE n.activity <> 'login';
GO

PRINT 'Created VW_UserSessionHistory';
GO

PRINT '';
PRINT '============================================================';
PRINT 'PEPS BOM Activity Tables — setup complete';
PRINT '============================================================';
