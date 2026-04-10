#  Backup 

@app.route('/dashboard')
def dashboard():
    if  os.path.exists(DATA_DIR):
        files = [f for f in os.listdir(DATA_DIR)
                    if os.path.isfile(os.path.join(DATA_DIR, f))]
        data = []
        patients_data = []
        
        
        for file in files:
            file_path = os.path.join(DATA_DIR,file)
            wb_data = load_workbook(file_path)
            ws_data = wb_data.active
            
            print("file:",file_path)
            for row in ws_data.iter_rows(min_row=2,values_only=True):
                print("row[6]:",row[6])
                if(row[6] == "Pending"):
                    data.append(list(row))
                elif(row[6] == "Accepted"):
                    patients_data.append(list(row))
            print("data in dahboard:",data)

        now = datetime.now()
        current_month = datetime.now().strftime("%m")
        current_year = now.year
        current_file = os.path.join(DATA_DIR,f"{current_month}-{current_year}.xlsx")
        print("current_file:",current_file)
        if  os.path.exists(current_file):
            wb_patients = load_workbook(current_file)
            ws_patients = wb_patients.active
            total_appointments = ws_patients.max_row - 1
            patients = 0
            for row in ws_patients.iter_rows(min_row=2,values_only=True):
                if(row[6] == 'Accepted'):
                    patients+=1
            
            overall_patients = get_overall_patients()
            conversion_rate = int((patients/total_appointments)*100)

            deviation = calculate_mom_for_month(int(current_month),int(current_year))
            print("result:",deviation)

            return render_template("Dashboard.html",pending_data = data,patients_data=patients_data,total_appointments = total_appointments,patients=patients,
                                   conversion_rate=conversion_rate,overall_patients=overall_patients,deviation=deviation)
        else:
          print("File not exists!!")  

    else:
        print("Data directory not exists!!")
    return render_template("Dashboard.html")     