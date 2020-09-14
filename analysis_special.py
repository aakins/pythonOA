import excel
import analysis
import pandas as pd

def duplicate_xfmr(idcolumn, idcolumn1, table):
    import pandas as pd
    c = "Error."
    facility_id = False
    global cell_style
    cell_style = 'Normal'
    global cell_alignment
    cell_alignment = 'False'
    #rows = excel.cursor.columns(table=table)
    for row in excel.cursor.columns(table=table):
        if "gs_facility_id" in row:
            facility_id = True
        else:
            facility_id = False
    if int(analysis.sumrows(table)) == 0:
        cell_style = 'Bad'
        c = "This table is empty."
    else:
        for row in excel.cursor.columns(table=table):
            if idcolumn in row:
                for row in excel.cursor.columns(table=table):
                    if idcolumn1 in row:
                        r = r""
                        sqlcommand = r.join(("SELECT {table}.OBJECTID, ".format(table=table)))
                        if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                            sqlcommand = r.join((sqlcommand, "{table}.gs_facility_id, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                        else:
                            sqlcommand = r.join((sqlcommand, "{table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                        if not ((idcolumn == "gs_phase") or (idcolumn1 == "gs_phase")):
                            sqlcommand = r.join((sqlcommand, ",{table}.gs_phase FROM {table} WHERE ((({table}.{idcolumn}) In (SELECT {idcolumn} FROM {table} As Tmp GROUP BY {idcolumn}, {idcolumn1}, gs_phase HAVING Count(*)>1 ))) GROUP BY {table}.OBJECTID, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn, idcolumn1=idcolumn1, table=table)))
                            if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                        else:
                            sqlcommand = r.join((sqlcommand, " FROM {table} WHERE ((({table}.{idcolumn}) In (SELECT {idcolumn} FROM {table} As Tmp GROUP BY {idcolumn}, {idcolumn1} HAVING Count(*)>1 ))) GROUP BY {table}.OBJECTID, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn, idcolumn1=idcolumn1, table=table)))
                            if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                        if not ((idcolumn == "gs_phase") or (idcolumn1 == "gs_phase")):
                            sqlcommand = r.join((sqlcommand, ", {table}.gs_phase ORDER BY {table}.{idcolumn}, {table}.OBJECTID, {table}.{idcolumn1}, {table}.gs_phase".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                            if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                        else:
                            sqlcommand = r.join((sqlcommand," ORDER BY {table}.{idcolumn}, {table}.OBJECTID, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                            if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                        excel.cursor.execute(sqlcommand)
                        row = excel.cursor.fetchall()
                        if not row:
                            cell_style = 'Normal'
                            cell_alignment = 'False'
                            c = "No duplicates found."
                        else:
                            cell_style = 'Bad'
                            c = str(len(row)) + " duplicates exist."
                            r = r""
                            sqlcommand = r.join(("SELECT {table}.OBJECTID, ".format(table=table)))
                            if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                sqlcommand = r.join((sqlcommand, "{table}.gs_facility_id, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                            else:
                                sqlcommand = r.join((sqlcommand, "{table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                            if not ((idcolumn == "gs_phase") or (idcolumn1 == "gs_phase")):
                                sqlcommand = r.join((sqlcommand, ",{table}.gs_phase FROM {table} WHERE ((({table}.{idcolumn}) In (SELECT {idcolumn} FROM {table} As Tmp GROUP BY {idcolumn}, {idcolumn1}, gs_phase HAVING Count(*)>1 ))) GROUP BY {table}.OBJECTID, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn, idcolumn1=idcolumn1, table=table)))
                                if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                    sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                            else:
                                sqlcommand = r.join((sqlcommand, " FROM {table} WHERE ((({table}.{idcolumn}) In (SELECT {idcolumn} FROM {table} As Tmp GROUP BY {idcolumn}, {idcolumn1} HAVING Count(*)>1 ))) GROUP BY {table}.OBJECTID, {table}.{idcolumn}, {table}.{idcolumn1}".format(idcolumn=idcolumn, idcolumn1=idcolumn1, table=table)))
                                if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                    sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                            if not ((idcolumn == "gs_phase") or (idcolumn1 == "gs_phase")):
                                sqlcommand = r.join((sqlcommand, ", {table}.gs_phase ORDER BY {table}.{idcolumn}, {table}.OBJECTID, {table}.{idcolumn1}, {table}.gs_phase".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                                if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                    sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                            else:
                                sqlcommand = r.join((sqlcommand," ORDER BY {table}.{idcolumn}, {table}.OBJECTID, {table}.{idcolumn1}".format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)))
                                if (facility_id == True and (not(idcolumn == "gs_facility_id" or idcolumn1 == "gs_facility_id"))):
                                    sqlcommand = r.join((sqlcommand, ", {table}.gs_facility_id".format(table=table)))
                            df = pd.read_sql_query(sqlcommand, excel.conn)
                            sheetname = excel.category + "-" + idcolumn[len("gs_"):]
                            analysis.createsheet(sheetname)
                            analysis.writedf(df, sheetname)
                            analysis.write_hyperlink(sheetname)
                        break
                    else:
                        cell_style = 'Neutral'
                        c = "The field ''" + str(idcolumn1) + "' needs to be added to the database."
            else:
                cell_style = 'Neutral'
                c = "The field ''" + str(idcolumn) + "' needs to be added to the database."
    return c

def duplicate_system():
    import pandas as pd
    c = "Error."
    global cell_style
    cell_style = 'Normal'
    global cell_alignment
    cell_alignment = 'False'
    table = "gs_service_point"
    #rows = excel.cursor.columns(table=table)
    if int(analysis.sumrows(table)) == 0:
        cell_style = 'Bad'
        c = "This table is empty."
    else:
        r = r""
        sqlcommand = r.join((r"""SELECT t.id AS 'ID',COUNT(T.ID) AS 'Count'
        FROM (SELECT gs_service_map_location AS 'ID' FROM GS_SERVICE_POINT"""))
        table = "gs_transformer"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT gs_equipment_location AS 'ID' FROM gs_transformer
                WHERE (((gs_transformer.gs_equipment_location) 
                    In (SELECT gs_equipment_location 
                    FROM gs_transformer As Tmp 
                    GROUP BY gs_equipment_location, gs_bank_id, gs_phase HAVING Count(*)>1 ))) """))
        table = "gs_overcurrent_device"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_OVERCURRENT_DEVICE"""))
        table = "gs_switch"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_SWITCH"""))
        table = "gs_capacitor_bank"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_CAPACITOR_BANK"""))
        table = "gs_generator"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_GENERATOR"""))
        table = "gs_motor"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_MOTOR"""))
        # table = "gs_street_light"
        # for row in excel.cursor.columns(table=table):
            # if "gs_equipment_location" in row:
                # sqlcommand = r.join((sqlcommand, """
                # UNION ALL
                # SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_STREET_LIGHT"""))
        table = "gs_voltage_regulator"
        for row in excel.cursor.columns(table=table):
            if "gs_equipment_location" in row:
                sqlcommand = r.join((sqlcommand, """
                UNION ALL
                SELECT GS_EQUIPMENT_LOCATION AS 'ID' FROM GS_VOLTAGE_REGULATOR"""))
        sqlcommand = r.join((sqlcommand, """) AS T group by t.id having count(*)>1"""))
        excel.cursor.execute(sqlcommand)
        row = excel.cursor.fetchall()
        if not row:
            cell_style = 'Normal'
            cell_alignment = 'False'
            c = "No duplicates found."
        else:
            cell_style = 'Bad'
            c = str(len(row)) + " duplicates exist."
            df = pd.read_sql_query(sqlcommand, excel.conn)
            sheetname = excel.category + "-SystemDuplicates"
            analysis.createsheet(sheetname)
            analysis.writedf(df, sheetname)
            analysis.write_hyperlink(sheetname)
    return c

def xfmr_voltage(idcolumn, idcolumn1, table):
    import pandas as pd
    #c = "Error."
    #facility_id = False
    #global cell_style
    excel.cell_style = 'Normal'
    #global cell_alignment
    excel.cell_alignment = 'False'
    rows = excel.cursor.columns(table=table)
    for row in rows:
        if idcolumn in row:
            for row in excel.cursor.columns(table=table):
                if idcolumn1 in row:
                    #r = r""
                    sqlcommand = r"""SELECT {table}.{idcolumn}, {table}.{idcolumn1}, Count({table}.OBJECTID) AS [Count]
                    FROM {table}
                    GROUP BY {table}.{idcolumn}, {table}.{idcolumn1}
                    ORDER BY Count({table}.OBJECTID) DESC
                    """.format(idcolumn=idcolumn,idcolumn1=idcolumn1,table=table)
                    excel.cursor.execute(sqlcommand)
                    row = excel.cursor.fetchall()
                    df = pd.read_sql_query(sqlcommand, excel.conn)
                    sheetname = excel.category + "-VoltageTable"
                    analysis.createsheet(sheetname)
                    analysis.writedf(df, sheetname)
                    break
    return
    
def span_assembly(table):
    import pandas as pd
    excel.c = "Error."
    excel.facility_id = False
    #global cell_style
    excel.cell_style = 'Normal'
    #global cell_alignment
    excel.cell_alignment = 'False'
    rows = excel.cursor.columns(table=table)
    for row in rows:
        if excel.idcolumn in row:
            for row in excel.cursor.columns(table=table):
                if excel.idcolumn1 in row:
                    #r = r""
                    sqlcommand = r"""SELECT gs_assembly_ref.gs_assembly_name, gs_assembly_ref.gs_engineering_analysis_name, Count({table}.gs_guid) AS [Count], {table}.gs_subtype_cd, gs_assembly_ref.gs_assembly_description
                    FROM ({table} INNER JOIN gs_attached_assemblies ON {table}.gs_guid = gs_attached_assemblies.gs_network_feature_guid) INNER JOIN gs_assembly_ref ON gs_attached_assemblies.gs_assembly_guid = gs_assembly_ref.gs_guid
                    GROUP BY gs_assembly_ref.gs_assembly_name, gs_assembly_ref.gs_engineering_analysis_name, {table}.gs_subtype_cd, gs_assembly_ref.gs_assembly_description, gs_assembly_ref.gs_is_flow_assembly
                    HAVING (((gs_assembly_ref.gs_is_flow_assembly)='true' Or (gs_assembly_ref.gs_is_flow_assembly)='true'))
                    ORDER BY gs_assembly_ref.gs_assembly_name, gs_assembly_ref.gs_engineering_analysis_name, Count({table}.gs_guid) DESC
                    """.format(idcolumn=excel.idcolumn,idcolumn1=excel.idcolumn1,table=table)
                    excel.cursor.execute(sqlcommand)
                    row = excel.cursor.fetchall()
                    df = pd.read_sql_query(sqlcommand, excel.conn)
                    sheetname = excel.category + "-AssemblyTable"
                    analysis.createsheet(sheetname)
                    analysis.writedf(df, sheetname)
                    df = df[df.duplicated(subset=['gs_assembly_name'], keep=False)]
                    sheetname = excel.category + "-DupAssemblyTable"
                    analysis.createsheet(sheetname)
                    analysis.writedf(df, sheetname)
                    analysis.write_hyperlink(sheetname)
                    break
    return

def two_bank_xfmr(table):
    excel.c = "Error."
    excel.facility_id = False
    #global cell_style
    excel.cell_style = 'Normal'
    #global cell_alignment
    excel.cell_alignment = 'False'
    #r = r""
    sqlcommand = r"""DROP TABLE IF EXISTS #QueryTable"""
    excel.cursor.execute(sqlcommand)
    #r = r""
    sqlcommand = r"""SELECT
            s.OBJECTID AS StationObjectID
            ,s.gs_guid AS StationGuid
            ,t.OBJECTID AS TransformerObjectID
            ,t.gs_guid AS TransformerGuid
            ,t.gs_phase AS TransformerPhase
            ,t.gs_equipment_location AS TransformerEquipLoc
            ,t.gs_bank_id AS TransformerBankID
            ,t.gs_winding_connection AS TransformerWindingConnection
            INTO #QueryTable
            FROM GS_TRANSFORMER t
        JOIN GS_ATTACHED_ASSEMBLIES aa ON aa.gs_display_feature_guid = t.gs_guid
        JOIN gs_station s ON s.gs_guid = aa.gs_network_feature_guid
        WHERE s.gs_guid IN
            (
            SELECT aa.gs_network_feature_guid FROM gs_attached_assemblies aa
            JOIN GS_TRANSFORMER t
                ON t.gs_guid = aa.gs_display_feature_guid
            GROUP BY aa.gs_network_feature_guid
            HAVING COUNT(aa.gs_network_feature_guid) = 2
            )"""
    excel.cursor.execute(sqlcommand)
    #r = r""
    sqlcommand = r"""SELECT StationObjectID, TransformerObjectID, TransformerPhase, 
        TransformerEquipLoc, TransformerBankID, TransformerWindingConnection FROM #QueryTable ORDER BY StationObjectID"""
    #excel.cursor.execute(sqlcommand)
    #row = excel.cursor.fetchall()
    df = pd.read_sql_query(sqlcommand, excel.conn)
    sheetname = excel.category + "-2bank"
    analysis.createsheet(sheetname)
    analysis.writedf(df, sheetname)
    analysis.write_hyperlink(sheetname)
    return

def two_bank_diff_id(table):
    #global cell_style
    excel.cell_style = 'Normal'
    #rows = excel.cursor.columns(table=table)
    #r = r""
    sqlcommand = r"""SELECT
            StationObjectID
            ,TransformerEquipLoc
            ,TransformerBankID
            ,COUNT(*) AS 'Count'
            FROM #QueryTable
            GROUP BY StationObjectID, TransformerEquipLoc, TransformerBankID
            HAVING COUNT(TransformerBankID) <= 1"""
    #excel.cursor.execute(sqlcommand)
    #row = excel.cursor.fetchall()
    df = pd.read_sql_query(sqlcommand, excel.conn)
    countid = df['Count'].sum()
    if (countid > 0):
        excel.cell_style = 'Bad'
        c = str(countid) + " errors."
    sheetname = excel.category + "-2bank_diff_id"
    analysis.createsheet(sheetname)
    analysis.writedf(df, sheetname)
    analysis.write_hyperlink(sheetname)
    return c

def two_bank_winding(table):
    #global cell_style
    excel.cell_style = 'Normal'
    #rows = excel.cursor.columns(table=table)
    #r = r""
    sqlcommand = r"""SELECT
            StationObjectID
            ,TransformerObjectID
            ,TransformerPhase
            ,TransformerEquipLoc
            ,TransformerBankID
            ,TransformerWindingConnection
            FROM #QueryTable
            WHERE TransformerWindingConnection <> 5 OR TransformerWindingConnection IS NULL"""
    excel.cursor.execute(sqlcommand)
    row = excel.cursor.fetchall()
    df = pd.read_sql_query(sqlcommand, excel.conn)
    countid = len(row)
    if (countid > 0):
        excel.cell_style = 'Bad'
        c = str(countid) + " errors."
    sheetname = excel.category + "-2bank_winding"
    analysis.createsheet(sheetname)
    analysis.writedf(df, sheetname)
    analysis.write_hyperlink(sheetname)
    return c

def two_bank_phasing(table):
    #global cell_style
    excel.cell_style = 'Normal'
    #rows = excel.cursor.columns(table=table)
    #r = r""
    sqlcommand = r"""SELECT
            StationObjectID
            ,TransformerObjectID
            ,TransformerPhase
            ,TransformerEquipLoc
            ,TransformerBankID
            ,TransformerWindingConnection
            FROM #QueryTable"""
    excel.cursor.execute(sqlcommand)
    row = excel.cursor.fetchall()
    df = pd.read_sql_query(sqlcommand, excel.conn)
    grouped_df = df.groupby("TransformerEquipLoc")
    grouped_list = grouped_df["TransformerPhase"].agg(lambda column: "".join(column))
    grouped_list = grouped_list.reset_index(name="TransformerPhase")
    grouped_list = grouped_list[grouped_list.TransformerPhase != 'ABC']
    countid = len(row)
    if (countid > 0):
        excel.cell_style = 'Bad'
        c = str(countid) + " errors."
    sheetname = excel.category + "-2bank_phasing"
    analysis.createsheet(sheetname)
    analysis.writedf(grouped_list, sheetname)
    analysis.write_hyperlink(sheetname)
    return c

def neutral(idcolumn, table):
    import pandas as pd
    c = "Error."
    global cell_style 
    cell_style = 'Normal'
    global cell_alignment
    cell_alignment = 'False'
    for row in excel.cursor.columns(table=table):
        if idcolumn in row:
            file_list = []

            sqlcommand = r"""SELECT {table}.{idcolumn}, {table}.gs_subtype_cd, Count({table}.OBJECTID) AS [Count]
                        FROM {table}
                        WHERE {table}.gs_subtype_cd =  1
                        AND ({table}.{idcolumn} IS null OR {table}.{idcolumn} = '' OR {table}.{idcolumn} LIKE '%UNK%')
                        GROUP BY {table}.{idcolumn}, {table}.gs_subtype_cd
                        ORDER BY {table}.{idcolumn}, {table}.gs_subtype_cd, Count({table}.OBJECTID) DESC
                        """.format(idcolumn=idcolumn,table=table)
            excel.cursor.execute(sqlcommand)
            row = excel.cursor.fetchall()
            totalrows = analysis.sumrows(table)
            
            try:
                if int(analysis.sumrows(table)) == int(row[0][2]):
                    cell_style = 'Normal'
                    cell_alignment = 'False'
                    if row[0][0] == None:
                        c = "All fields are NULL."
                    if row[0][0] == 0:
                        c = "All fields are populated with '0'."
                    else:
                        c = "All fields are populated with '" + str(row[0][0]) +"'."
                    if (row[0][0] == None) or (row[0][0] == "") or (row[0][0] == "UNK"):
                        cell_style = 'Bad'
                        sqlcommand = r"""SELECT {table}.OBJECTID, {table}.{idcolumn}
                                    FROM {table}
                                    WHERE {table}.gs_subtype_cd =  1
                                    AND ({table}.{idcolumn} IS null OR {table}.{idcolumn} = '' OR {table}.{idcolumn} LIKE '%UNK%')
                                    GROUP BY {table}.OBJECTID, {table}.{idcolumn}
                                    ORDER BY {table}.OBJECTID, {table}.{idcolumn}
                                    """.format(idcolumn=idcolumn,table=table)
                        df = pd.read_sql_query(sqlcommand, excel.conn)
                        sheetname = excel.category + "-" + idcolumn[len("gs_"):]
                        analysis.createsheet(sheetname)
                        analysis.writedf(df, sheetname)

                else:
                    for row in row:
                        if (row[0] == None) or (row[0] == "") or (row[0] == "UNK"):
                            cell_style = 'Bad'
                        fcol = row[0]
                        if fcol == 0:
                            fcol = "0"
                        if fcol == None:
                            fcol = "NULL"
                        fnum = row[2]
                        fper = int(round(int(fnum)/int(totalrows)*100))
                        file_list.append({"fcol": fcol, "fnum": fnum, "fper": fper})
                    cell_alignment = 'True'
                    c = r""        
                    c = c.join(("{fnum} ({fper}%) populated with '{fcol}'.\n".format(fcol=fl['fcol'], fnum=fl['fnum'], fper=fl['fper']) for fl in file_list))
                    if cell_style == 'Bad':
                        sqlcommand = r"""SELECT {table}.OBJECTID, {table}.{idcolumn}
                                    FROM {table}
                                    WHERE {table}.gs_subtype_cd =  1
                                    AND ({table}.{idcolumn} IS null OR {table}.{idcolumn} ='' OR {table}.{idcolumn} LIKE '%UNK%')
                                    GROUP BY {table}.OBJECTID, {table}.{idcolumn}
                                    ORDER BY {table}.OBJECTID, {table}.{idcolumn}
                                    """.format(idcolumn=idcolumn,table=table)
                        df = pd.read_sql_query(sqlcommand, excel.conn)
                        sheetname = excel.category + "-" + idcolumn[len("gs_"):]
                        analysis.createsheet(sheetname)
                        analysis.writedf(df, sheetname)
                        analysis.write_hyperlink(sheetname)
                break
            except IndexError:
                break
        else:
            cell_style = 'Neutral'
            c = "This field needs to be added to the database."
    
    return c

def is_feeder_bay(idcolumn, table):
    import pandas as pd
    c = "Error."
    facility_id = False
    for row in excel.cursor.columns(table=table):
        if "gs_facility_id" in row:
            facility_id = True
    global cell_style 
    cell_style = 'Normal'
    global cell_alignment
    cell_alignment = 'False'
    rows = excel.cursor.columns(table=table)
    if int(analysis.sumrows(table)) == 0:
                cell_style = 'Bad'
                c = "This table is empty."
    else:
        for row in excel.cursor.columns(table=table):
            if idcolumn in row:
                r = r""
                sqlcommand = r.join(("SELECT {table}.OBJECTID, {table}.{idcolumn}".format(idcolumn=idcolumn,table=table)))
                if (facility_id == True and (not(idcolumn == "gs_facility_id"))):
                    sqlcommand = r.join((sqlcommand, """,{table}.gs_facility_id, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay
                        FROM {table} 
                        WHERE ({table}.{idcolumn} = '4' AND ({table}.gs_is_feeder_bay <> '1' OR {table}.gs_is_feeder_bay IS NULL)) OR ({table}.gs_overcurrent_device_subtype <> '4' AND {table}.gs_is_feeder_bay <> '0') 
                        GROUP BY {table}.{idcolumn}, {table}.OBJECTID, {table}.gs_facility_id, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay""".format(idcolumn=idcolumn,table=table)))
                else:
                    sqlcommand = r.join((sqlcommand, """, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay
                        FROM {table} 
                        WHERE ({table}.{idcolumn} = '4' AND ({table}.gs_is_feeder_bay <> '1' OR {table}.gs_is_feeder_bay IS NULL)) OR ({table}.gs_overcurrent_device_subtype <> '4' AND {table}.gs_is_feeder_bay <> '0') 
                        GROUP BY {table}.{idcolumn}, {table}.OBJECTID, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay""".format(idcolumn=idcolumn,table=table)))
                excel.cursor.execute(sqlcommand)
                rows = excel.cursor.fetchall()
                totalrows = analysis.sumrows(table)
                if not rows:
                    cell_style = 'Normal'
                    cell_alignment = 'False'
                    c = "All are populated correctly."
                else:
                    percent = int(round(int(len(rows))/int(totalrows)*100))
                    cell_style = 'Bad'
                    c = str(len(rows)) + " (" + str(percent) + "%)" + " incorrect values."
                    r = r""
                    sqlcommand = r.join(("SELECT {table}.OBJECTID, {table}.{idcolumn}".format(idcolumn=idcolumn,table=table)))
                    if (facility_id == True and (not(idcolumn == "gs_facility_id"))):
                        sqlcommand = r.join((sqlcommand, """,{table}.gs_facility_id, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay
                            FROM {table} 
                            WHERE ({table}.{idcolumn} = '4' AND ({table}.gs_is_feeder_bay <> '1' OR {table}.gs_is_feeder_bay IS NULL)) OR ({table}.gs_overcurrent_device_subtype <> '4' AND {table}.gs_is_feeder_bay <> '0') 
                            GROUP BY {table}.{idcolumn}, {table}.OBJECTID, {table}.gs_facility_id, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay""".format(idcolumn=idcolumn,table=table)))
                    else:
                        sqlcommand = r.join((sqlcommand, """, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay
                            FROM {table} 
                            WHERE ({table}.{idcolumn} = '4' AND ({table}.gs_is_feeder_bay <> '1' OR {table}.gs_is_feeder_bay IS NULL)) OR ({table}.gs_overcurrent_device_subtype <> '4' AND {table}.gs_is_feeder_bay <> '0') 
                            GROUP BY {table}.{idcolumn}, {table}.OBJECTID, {table}.gs_overcurrent_device_subtype, {table}.gs_is_feeder_bay""".format(idcolumn=idcolumn,table=table)))
                    df = pd.read_sql_query(sqlcommand, excel.conn)
                    sheetname = excel.category + "-FeederBay"
                    analysis.createsheet(sheetname)
                    analysis.writedf(df, sheetname)
                    analysis.write_hyperlink(sheetname)
                break
            else:
                cell_style = 'Neutral'
                c = "This field needs to be added to the database."  
    return c