from flask import Flask, request, send_file, render_template
import pandas as pd
import io
import os
import numpy as np
from ortools.linear_solver import pywraplp

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def truckLoader():
    if request.method == 'POST':
        # Get the uploaded file
        file = request.files.get('file')
        if file:
            # Read Excel file into a DataFrame
            shipments_data = pd.read_excel(file, sheet_name = 0)
            truck_data = pd.read_excel(file, sheet_name = 1)
            
            # --- Process the DataFrame ---
            # For example, add a new column named 'Processed' with a default value.
            truckwise_shipments, shipmentwise_trucks = solve(file)

            # ------------------------------
            # Save the modified DataFrame to a BytesIO stream as an Excel file.
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                shipmentwise_trucks.to_excel(writer, index=False, sheet_name='shipmentwise_trucks')
                truckwise_shipments.to_excel(writer, index=False, sheet_name='truckwise_shipments')
            output.seek(0)
            
            # Return the modified file as an attachment for download.
            return send_file(
                output,
                as_attachment=True,
                download_name='Optimal Loading plan.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    # Render the HTML form for GET requests.
    return render_template('Truckloader.html')

def create_truckloader_data(file):
  shipments = pd.read_excel(file, sheet_name = 0)
  trucks = pd.read_excel(file, sheet_name = 1)

  data = {}
  data["shipments"] = shipments
  data["trucks"] = trucks
  data["shipmentsNumber"] = list(range(shipments.shape[0]))
  data["trucktypesNumber"] = list(range(trucks.shape[0]))
  return data

def solve(file):
    df = create_truckloader_data(file)
    
    # Create the mip solver with the SCIP backend.
    solver = pywraplp.Solver.CreateSolver("SCIP")

    if not solver:
        return

    # Variables
    # x[i, j] = 1 if shipment i is filled in truck type j and truck number k.
    x = {}
    for i in df["shipmentsNumber"]:
        for j in df["trucktypesNumber"]:
            for k in list(range(df["trucks"].loc[j, "Number of Trucks"])):
                if(df["shipments"].loc[i, "Origin"] == df["trucks"].loc[j, "Origin"] and df["shipments"].loc[i, "Destination"] == df["trucks"].loc[j, "Destination"]):
                    x[(i, j, k)] = solver.IntVar(0, 1, "x_%i_%i_%i" % (i, j, k))
                else:
                    x[(i, j, k)] = solver.IntVar(0, 0, "x_%i_%i_%i" % (i, j, k))
                      
    # y[(j, k)] = 1 if truck type j and truck number k is used.
    y = {}
    for j in df["trucktypesNumber"]:
      for k in list(range(df["trucks"].loc[j, "Number of Trucks"])):
        y[(j, k)] = solver.IntVar(0, 1, "y[%i, %i]" % (j, k))

    # Constraints
    # Each shipment should go in exactly one truck.
    for i in df["shipmentsNumber"]:
        solver.Add(sum(x[i, j, k] for j in df["trucktypesNumber"] for k in list(range(df["trucks"].loc[j, "Number of Trucks"]))) == 1)

    # The shipments packed in each truck cannot exceed its weight capacity.
    for j in df["trucktypesNumber"]:
      for k in list(range(df["trucks"].loc[j, "Number of Trucks"])):
        solver.Add(
            sum(x[(i, j, k)] * df["shipments"].loc[i, "Weight"] for i in df["shipmentsNumber"])
            <= y[(j, k)] * df["trucks"].loc[j, "Truck Capacity (Kg Weight)"]
        )

    # The shipments packed in each truck cannot exceed its volumetric capacity.
    for j in df["trucktypesNumber"]:
      for k in list(range(df["trucks"].loc[j, "Number of Trucks"])):
        solver.Add(
            sum(x[(i, j, k)] * df["shipments"].loc[i, "Volume"] for i in df["shipmentsNumber"])
            <= y[(j, k)] * df["trucks"].loc[j, "Truck Capacity (Cubic Meter Volume)"])

    # Objective: minimize the number of bins used.
    solver.Minimize(solver.Sum([y[(j, k)] for j in df["trucktypesNumber"] for k in list(range(df["trucks"].loc[j, "Number of Trucks"]))]))
    
    print("Finding the best loading plan")
    status = solver.Solve()

    output = pd.DataFrame(columns=['Truck','Origin', 'Destination', 'Shipments'])
    df["shipments"]["Truck"] = ""

    if status == pywraplp.Solver.OPTIMAL:        
        num_trucks = 0
        for j in df["trucktypesNumber"]:
           for k in list(range(df["trucks"].loc[j, "Number of Trucks"])):
            if y[(j, k)].solution_value() == 1:
                truck_shipments = []
                truck_weight = 0
                truck_volume = 0
                for i in df["shipmentsNumber"]:
                    if x[i, j, k].solution_value() > 0:
                        truck_shipments.append(i+1)
                        truck_weight += df["shipments"].loc[i, "Weight"]
                        truck_volume += df["shipments"].loc[i, "Volume"]
                        df["shipments"].loc[i, 'Truck'] = str(j +1) + "_" + str(k+1)
                        
                if truck_shipments:
                    num_trucks += 1
                    output = pd.concat([output, pd.DataFrame({'Truck': str(j +1) + "_" + str(k+1), 'Origin': df["trucks"].loc[j, "Origin"], 'Destination': df["trucks"].loc[j, "Destination"], 'Shipments': [truck_shipments]})], ignore_index=True)
                    '''print("Truck number", j+1, k+1)
                    print("Shipments packed:", truck_shipments)
                    print("Total weight:", truck_weight)
                    print("Total volume", truck_volume)
                    print()'''
        print()
        print("Number of trucks used:", num_trucks)
        print("Time = ", solver.WallTime()/1000, "Seconds")

        # Write output to csv files
        #output.to_csv('C:/Users/shrinath/TruckLoader/Python App/Truckwise shipments.csv')
        #df["shipments"].to_csv('C:/Users/shrinath/TruckLoader/Python App/Shipmentwise trucks.csv', index = False)
    else:
        print("The problem does not have an optimal solution. Check the number of available trucks.")
    return output, df["shipments"]

if __name__ == '__main__':
    app.run(debug=True)
