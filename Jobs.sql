select DISTINCT
                        t.NAME, j.NAME as JOB, CAST(ja_d.VALUE_STR AS VARCHAR(255)) as DIR_PATH, d.DIRECTORY_NAME
                        , db.ID_DATABASE, db.NAME, db.ID_TYPE, db.ID_CONTYPE, db.HOST_NAME, CAST(db.DATABASE_NAME AS VARCHAR(255)) DATABASE_NAME, db.PORT, db.USERNAME, db.PASSWORD, db.SERVERNAME, db.DATA_TBS, db.INDEX
                        --, d2.*
                        from R_T t (nolock)
                        join R_JOBEN ja (nolock) on (ja.CODE = 'name' and cast(ja.VALUE_STR as varchar(50)) = t.NAME)
                        join R_JOBEB_ATT ja_d (nolock) on (ja.ID_JOBEN = ja_d.ID_JOBEN and ja_d.CODE = 'dir_path')
                        join R_JOB j (nolock) on (j.ID_JOB = ja.ID_JOB)
                        join R_DIRECTORY d (nolock) on (d.ID_DIRECTORY = j.ID_DIRECTORY)
                        join R_DATABASE st (nolock) on (st.ID_TRANSF = t.ID_TRANSF)
                        join R_DATABASE db (nolock) on (db.ID_DATABASE = st.ID_DATABASE)
                        JOIN R_DIRECTORY d2 (nolock) ON t.ID_DIRECTORY = d.ID_DIRECTORY
                        --WHERE j.NAME = 'Job_Load'
                        --AND d.DIRECTORY_NAME = 'CB_R'
                        --ORDER BY 1, 4
                        UNION ALL
                        SELECT T3.NAME, T1.NAME, NULL, NULL, T5.*
                        from R_JOB as T1
                        JOIN R_JOB_HOP as T2 ON T1.ID_JOB = T2.ID_JOB
                        JOIN R_JOB as T3 ON T2.ID_JOB_COPY_TO = T3.ID_JOB
                        JOIN R_JOB_DATABASE as T4 ON T3.ID_JOB = T4.ID_JOB
                        JOIN R_DATABASE as T5 ON T4.ID_DATABASE = T5.ID_DATABASE
                        ORDER BY 2, 1;
