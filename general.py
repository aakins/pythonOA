ALLOWED_EXTENSIONS = {'gz'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def analyze_file(filename):
    import os
    from datetime import date
    import calendar
    import re
    import openpyxl
    import gzip,shutil
    month = calendar.month_abbr[date.today().month]
    year = date.today().year
    day = date.today().day
    #username = os.environ["USERNAME"]
    #onedrivedir = "C:/Users/"+ username +"/OneDrive - National Information Solutions Cooperative/"
    #downloadsdir = "C:/Users/"+ username +"/Downloads/"
    projectsdir = "C:/Projects/Uploads/"
    #coopsdir = r"\\mofs\sdesvr\Support\MDM-DA\Coops"
    # companydir = str(number) + " " + name + "/"
    # monthdir = month + " " + str(year) + "/"
    global number
    number = re.search('([0-9])+', filename)
    number = number.group(0)

    #if not os.path.exists(onedrivedir + companydir):
    #    os.mkdir(onedrivedir + companydir)

    # if not os.path.exists(onedrivedir + companydir + monthdir):
        # os.mkdir(onedrivedir + companydir + monthdir)
        
    # if not os.path.exists(projectsdir + companydir):
        # os.mkdir(projectsdir + companydir)
        
    # if not os.path.exists(projectsdir + companydir + monthdir):
        # os.mkdir(projectsdir + companydir + monthdir)

    gzfilename, file_extension = os.path.splitext(filename)
    # if not os.path.exists(projectsdir + companydir + monthdir):
        # os.mkdir(projectsdir + companydir + monthdir)
        
    if file_extension == ".gz":
        with gzip.open(projectsdir + filename, 'r') as f_in, open(projectsdir + filename[:-3], 'wb') as f_out:
              shutil.copyfileobj(f_in, f_out)
    gzfile = projectsdir + gzfilename
    restore_db(gzfile, number, projectsdir)
    set_level(number)
    
    xltemplatefile = projectsdir + "OAMapWiseDataReview_Master_temp.xlsm"
    global xlanalysisfile
    xlanalysisfile = projectsdir + "gs" + str(number) + "-OAMapWiseDataReview-" + str(day) + str(month) + str(year) + ".xlsm"
    global xlanalysisfilename
    xlanalysisfilename = "gs" + str(number) + "-OAMapWiseDataReview-" + str(day) + str(month) + str(year) + ".xlsm"
    if os.path.exists(xltemplatefile):
        mywb = openpyxl.load_workbook(xltemplatefile, keep_vba=True)
        mywb.save(xlanalysisfile)
    else:
        print("File " + xltemplatefile + " not found.")
    
    
    return xlanalysisfilename, xlanalysisfile

def restore_db(gzfile, number, projectsdir):
    import pyodbc
    import os
    import time

    conn = pyodbc.connect("Driver={SQL Server};Server=.\SQLEXPRESS;Database=master;Trusted_Connection=yes", autocommit = True)
    conn.timeout = 60
    cursor = conn.cursor()
    file_list = []
    print("Restoring DB...")
    
    def get_filelistonly(bak_file):
        sqlcommand = r"""
                        RESTORE filelistonly FROM DISK = N'{bak_file}'
                     """.format(bak_file=bak_file)
        cursor.execute(sqlcommand)
        rows = cursor.fetchall()

        for row in rows:
            fname = row[0]
            fext = os.path.splitext(row[1])[1]
            if "." not in fext:
                raise ValueError("No extension found in row")
            file_list.append({"fname": fname, "fext": fext})
        return file_list
    
    def get_drop_command(new_db):
        r = None
        if len(file_list) > 0:
            sqlcommand = r"""DROP DATABASE IF EXISTS {new_db}
                         """.format(new_db=new_db)
            r = sqlcommand
            try:
                cursor.execute(sqlcommand)
                while cursor.nextset():
                    pass
            except:
                print("Couldn't drop table")
                pass
        return r
    
    def get_restore_command(new_db, bak_file, file_list):
        r = None
        if len(file_list) > 0:
            sqlcommand = r"""RESTORE DATABASE {new_db} FROM DISK = N'{bak_file}'
                            WITH
                            FILE = 1,
                         """.format(new_db=new_db, bak_file=bak_file)
            sqlcommand = sqlcommand + ", \n".join(("MOVE N'{fname}' TO N'{projectsdir}\{new_db}{fext}'".format(fname=fl['fname'], fext=fl['fext'], new_db=new_db, number=number, projectsdir=projectsdir) for fl in file_list))
            sqlcommand = sqlcommand + ", NOUNLOAD, REPLACE, STATS = 5"
            r = sqlcommand
            sqlcommand = sqlcommand.replace("/" , "\\")
            try:
                cursor.execute(sqlcommand)
                while cursor.nextset():
                    pass
            except:
                print("Couldn't restore table")
                pass
        return r

    #rows_empty = ()
    #backup_file = projectsdir + companydir + monthdir + gzfilename[:-3]
    file_list = get_filelistonly(gzfile)

    r = get_drop_command("gs" + number)
    time.sleep(15)
    r = get_restore_command("gs" + number, gzfile, file_list)
    time.sleep(15)
    conn.close()
    return

def set_level(number):
    import pyodbc

    conn = pyodbc.connect("Driver={SQL Server};Server=.\SQLEXPRESS;Database=master;Trusted_Connection=yes", autocommit = True)
    conn.timeout = 60
    cursor = conn.cursor()

    sqlcommand = r"""Use master
                    ALTER DATABASE gs{number}
                    SET COMPATIBILITY_LEVEL = 130;
                 """.format(number=number)
    cursor.execute(sqlcommand)
    while cursor.nextset():
        pass

    conn.close()
    return