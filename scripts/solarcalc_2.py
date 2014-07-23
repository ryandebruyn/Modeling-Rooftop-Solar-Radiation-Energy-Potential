import arcpy
import os, sys
import SSReport as REPORT

arcpy.env.overwriteOutput = 1

# calculate the potential solar energy of the building
def solarCalc(infeat, inbldg):

  # calculate the total solar radiation from solar radiation analysis for location
  SolSumWHm2 = 0.0
  with arcpy.da.SearchCursor(infeat, "*") as cursor:
    for row in cursor:
      for i in range(2,28):
        SolSumWHm2 += row[i]
        ##print("{},{}").format(i-2,row[i])

  #get area of proposed solar panel polygon
  with arcpy.da.SearchCursor(inbldg, "SHAPE@AREA") as cursor:
      for row in cursor:
          shpArea_m = row[0]

  # Calculate solar electricity potential from the sun of the total allowable area
  # 20% loss due to edge and chimneys.
  areaAllow_m = shpArea_m * 0.80
  SolarPKWH = (areaAllow_m * (SolSumWHm2 / 1000))

  # Cost to build largest possible system on defined area (AveCost = $9.00/watt)
  # Based on a system with 208 watt solar modules (3'x5') = 13.867 w/ft2.
  areaAllow_ft = areaAllow_m * 10.7639104
  Cost = areaAllow_ft * 13.867 * 9.00

  #Cost of system minus rebates and tax break ($3.50/watt, 30% tax)
  Rebate = (areaAllow_ft * 13.867 * 3.50)
  TaxCredit = (Cost - Rebate) * 0.3
  Actualcost = Cost - Rebate - TaxCredit


  #'Calculate Maximum system for rooftop - Based on a 208W solar module (3'x5')
  #'SysSize (KW) = Area * 208W/15ft2 * 1KW/1000W
  SysSize = (areaAllow_ft * 13.867) / 1000

  #'Calculate Total System potential  [units = watts]
  SolSys = areaAllow_ft  * 13.867

  ##SolPSys = (SolSys * 0.68) / 0.114              #Solar energy potential [KWH/yr] of rooftop system - loss factors (STC Rating 5% * temp 6% * dust 7% * wiring 5% * DC/AC 10%)
  SolPSys = ((SolSys * 5 * 365) / 1000) * 0.68     #Solar Energy for average 5 hours a day / year * loss factor
  ##SolPSys = (SolSys / 0.114) * 0.15

  ERate = 0.12
  NetRate = 0.06
  Demand = 7300
  if SolPSys - Demand < 0:
    Net = 0
  else:
    Net = SolPSys - Demand

  DemandSav = Demand * ERate
  NetSav = Net * NetRate
  TotalSav = DemandSav + NetSav
  Payback = Actualcost / TotalSav
  ProdPerc = (SolPSys / Demand * 100)

  return ["",areaAllow_ft,SysSize,Cost,Rebate,TaxCredit,Actualcost,
          "",SolarPKWH,SolSys,SolPSys,
          Demand,ProdPerc,DemandSav,NetSav,TotalSav,Payback]

def CreateTableReport(fileName,calc_values):
    pdfOutput = REPORT.openPDF(fileName)

    #Set up Table
    NumColumns = 4
    report = REPORT.startNewReport(NumColumns,
                                   title = 'Your Solar Energy Potential Report ',
                                   landscape = False,
                                   numRows = "", # probably empty
                                   titleFont = REPORT.ssTitleFont)
    grid = report.grid


    keys = ["PV System Investment ",
            "Area of rooftop (ft2)",
            "System size (Kw)",
            "Total Cost of System ($)",
            "      Rebates ($3.50/W)",
            "      Tax Credit (30%)",
            "Actual Cost ($)",
            "PV System & Savings ",
            "Total Annual Solar Energy (KwHm2/yr)",
            "System Peak Power Output (watts)",
            "System Annual Energy Output(KwH/yr)",
            "Average HouseHold Use (KwH/yr)",
            "Production Percentage",
            "Saving from usage ($/yr)",
            "Saving from net metering ($yr)",
            "Total Savings",
            "Years to Pay Back System "]

    vals = calc_values

    for keyInd, keyVal in enumerate(keys):
        cVal = vals[keyInd]
        #write item key to column 2
        if "PV System" in keyVal:
          grid.stepRow() # move to next row
          grid.createLineRow(grid.rowCount, startCol = 1, endCol = 3)
          grid.stepRow() # move to next row
          grid.writeCell((grid.rowCount, 1),
                   "{}   ".format(keyVal),
                   justify = "left",
                   fontObj = REPORT.ssBoldFont,
                   color = "Maroon",
                   )
          grid.stepRow() # move to next row
          grid.createLineRow(grid.rowCount, startCol = 1, endCol = 3)
        else:
          grid.writeCell((grid.rowCount, 1),
                   "{}  =".format(keyVal),
                   justify = "left",
                   )

        # write item value to column 3
        if cVal:
          if "Total Savings" in keyVal:
            grid.writeCell((grid.rowCount, 2),
                       round(cVal,2),
                       justify = "right",
                       fontObj = REPORT.ssBoldFont
                       )
          elif "Cost" in keyVal:
            grid.writeCell((grid.rowCount, 2),
                       round(cVal,2),
                       justify = "right",
                       fontObj = REPORT.ssBoldFont
                       )
          else:
            grid.writeCell((grid.rowCount, 2),
                     round(cVal,2),
                     justify = "right",
                     )
        grid.stepRow() # move to next row

    grid.createLineRow(grid.rowCount, startCol = 1, endCol = 3)
    grid.stepRow()
    grid.finalizeTable() # Will fill empty rows with spaces.
    report.write(pdfOutput) # write to PDF
    pdfOutput.close()

    return fileName

def CreatePDFReport(inDirectRadFeat, inPVFeat, finalpdf_filename):
  #mxd = arcpy.mapping.MapDocument("CURRENT")
  mxd_filename = 'D:\\UC2012\\demos\\solar\\LArooftop\\SolarRoofDemo_UC2012_Final01.mxd'
  layoutpdf_filename = os.path.join(os.getcwd(),'Report1_Layout.pdf')
  tablepdf_filename = os.path.join(os.getcwd(),'Report2_TableReport.pdf')

  #Create layout report (page1)
  arcpy.AddMessage("Creating Layout...")
  mxd = arcpy.mapping.MapDocument(mxd_filename)
  #mxd = arcpy.mapping.MapDocument("CURRENT")
  arcpy.mapping.ExportToPDF(mxd,layoutpdf_filename)
  LayoutPDF = layoutpdf_filename

  #Create Table Report(page2)
  arcpy.AddMessage("Calculating Solar Values...")
  calc_values = solarCalc(inDirectRadFeat, inPVFeat)
  TablePDF = CreateTableReport(tablepdf_filename,calc_values)

  #Create Final report (merge)
  msg = ("Creating Output Report {}".format(finalpdf_filename))
  arcpy.AddMessage(msg)
  if os.path.exists(finalpdf_filename):
      os.remove(finalpdf_filename)
  pdfFinal = arcpy.mapping.PDFDocumentCreate(finalpdf_filename)
  pdfFinal.appendPages(LayoutPDF)
  pdfFinal.appendPages(TablePDF)
  pdfFinal.saveAndClose()

  return finalpdf_filename


if __name__ == "__main__":

  inDirectRadFeat = arcpy.GetParameterAsText(0)
  inPVFeat = arcpy.GetParameterAsText(1)
  FinalReportName = arcpy.GetParameterAsText(3)

  ##inDirectRadFeat = r"d:\UC2012\demos\solar\LArooftop\output\pnt_direct.shp"
  ##inPVFeat = r"d:\UC2012\demos\solar\LArooftop\output\fgdb.gdb\SolarBldg"
  ##FinalReportName = os.path.join(os.getcwd(),'FinalSolarReport.pdf')

  FinalReport = CreatePDFReport(inDirectRadFeat,inPVFeat,FinalReportName)

  print "Complete"