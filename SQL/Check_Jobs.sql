SELECT T1.NAME as Job_name, T3.NAME as Job_transform
                            from R_JOB as T1
                            JOIN R_JOB_HOP as T2 ON T1.ID_JOB = T2.ID_JOB
                            JOIN R_JOBEN as T3 ON T2.ID_JOBEN_COPY = T3.ID_JOBEN
                            WHERE T3.NAME in (select NAME from R_JOB)
