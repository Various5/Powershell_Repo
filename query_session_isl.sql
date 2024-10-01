copy (
SELECT
    s.dbid,
    s.d_clienthostname,
    s.d_createdtime,
    to_timestamp(s.d_createdtime) AT TIME ZONE 'UTC' AS created_time_readable,
    s.d_deskhostname,
    s.d_started_from_module,
    s.d_starttime,
    to_timestamp(s.d_starttime) AT TIME ZONE 'UTC' AS start_time_readable,
    -- Include last change times
    s.d_desk_last_change,
    to_timestamp(s.d_desk_last_change) AT TIME ZONE 'UTC' AS desk_last_change_readable,
    s.d_client_last_change,
    to_timestamp(s.d_client_last_change) AT TIME ZONE 'UTC' AS client_last_change_readable,
    -- Calculate end time with fallback
    COALESCE(
        GREATEST(s.d_desk_last_change, s.d_client_last_change, s.d_starttime)
    ) AS end_time,
    to_timestamp(
        COALESCE(
            GREATEST(s.d_desk_last_change, s.d_client_last_change, s.d_starttime)
        )
    ) AT TIME ZONE 'UTC' AS end_time_readable,
    -- Calculate session duration
    COALESCE(
        GREATEST(s.d_desk_last_change, s.d_client_last_change, s.d_starttime)
    ) - s.d_starttime AS duration_seconds,
    AGE(
        to_timestamp(
            COALESCE(
                GREATEST(s.d_desk_last_change, s.d_client_last_change, s.d_starttime)
            )
        ),
        to_timestamp(s.d_starttime)
    ) AS duration_readable,
    s.d_user,
    s.d_user_id,
    s.d_username,
    COALESCE(
        STRING_AGG(convert_from(m.d_value, 'UTF8'), '; '),
        'No messages'
    ) AS messages
FROM
    public.isllight_session_v s
LEFT JOIN
    public.isllight_session_qe_message_e m ON s.dbid = m.d_dbid
WHERE
    s.d_starttime >= EXTRACT(EPOCH FROM (NOW() AT TIME ZONE 'UTC')) - 86400  -- Last 24 hours
GROUP BY
    s.dbid,
    s.d_clienthostname,
    s.d_createdtime,
    s.d_deskhostname,
    s.d_started_from_module,
    s.d_starttime,
    s.d_desk_last_change,
    s.d_client_last_change,
    s.d_user,
    s.d_user_id,
    s.d_username
) TO STDOUT WITH CSV HEADER;
