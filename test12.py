import openpyxl

#openpyxl works with xlsx only, for old xls file we can use xlrd 
wb_obj = openpyxl.load_workbook('test1.xlsx')
sheet = wb_obj.active

exceldata = []
for row in sheet.values:
    exceldata.append(row[2])


all_data = []
gain= []
loss= []
class DataReader:
    '''This class takes a list of column data as input and gives gain,
     loss and avg gain and avg loss these results can be written to excel file as well'''
    def test(self, data):
            for i in range(len(data)):
                Data ={"inp":0, "change":0, "gain":0, "loss":0, "Avggain":0, "Avgloss":0, "MF":0}
                if i==0:
                    Data["inp"] = round(data[i],3)   
                if i > 0:
                    Data["inp"] = round(data[i],3)
                    Data["change"] = round(data[i] - data[i-1],3)
                    if Data["change"]>=0:
                        Data["gain"] = abs(Data["change"])
                        gain.append(Data["gain"])
                    if Data["change"]<=0:
                        Data["loss"]  = abs(Data["change"])
                        loss.append(Data["loss"])
                    if i==14:
                        Data["Avggain"] = round(sum(gain)/14,3)
                        Data["Avgloss"] = round(sum(loss)/14,3)
                    if i>14:
                        Data["Avggain"] = round(((all_data[i-1]["Avggain"]*13)+Data["gain"])/14,3)
                        Data["Avgloss"] = round(((all_data[i-1]["Avgloss"]*13)+Data["loss"])/14,3)
                    try:
                        Data["MF"] = round(Data["Avggain"]/Data["Avgloss"],3)
                    except:
                        Data["MF"] = 0

                all_data.append(Data)
            return all_data

data_obj = DataReader()
data_obj.test(exceldata)
print(all_data)
