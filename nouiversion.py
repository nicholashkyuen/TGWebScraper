import xlwings as xw
import waybilltoexcel as funcs
from tkinter.messagebox import showerror


def no_info_error(forwarder,waybillno):
    showerror("Warning", "Error: [{}] waybill - {} does not exist / contains no information, please retry.".format(forwarder,waybillno))

wb = xw.books.active

selection = wb.selection.options(ndim=2, index=False, convert=str).columns[2:10]

print(selection[0])

for area in selection:
    print(area[4].options(numbers=int).value)

raw = []
forwarder = ["DHL",'FEDEX','SF','TNT','UPS']

for area in selection:
    if area[0].value is not None and area[4].value is not None and area[0].value.upper() in forwarder:
        item = [area[0].value,area[4].options(numbers=int).value]
        raw.append(item)

    else:
        raw.append(None)
        

input_data =  raw
    
engine = funcs.setup(input_data)

for thing in input_data:
    if thing is None:
        continue
    forwarder = thing[0]
    print(forwarder)
    if forwarder.lower() == 'ups':
        if funcs.case_ups(engine,thing[1]):
            no_info_error(thing[0],thing[1])
            funcs.result.append(None)
    elif forwarder.lower() == 'dhl':
        if funcs.case_dhl(engine,thing[1]):
            no_info_error(thing[0],thing[1])
            funcs.result.append(None)
    elif forwarder.lower() == 'sf':
        if funcs.case_sf(engine,thing[1]):
            no_info_error(thing[0],thing[1])
            funcs.result.append(None)
    else:
        if funcs.case_fedex(thing[0],thing[1]):
            no_info_error(thing[0],thing[1]) 
            funcs.result.append(None)
for line in input_data:
    if line is None:
        continue
    elif line[0] in ["UPS", "DHL", "SF"]:
        engine.quit()
        break

output = funcs.result

print(output)

a = 0

for iteration, j in enumerate(input_data):
    if j is not None and output[a] is not None:
        row = selection[iteration]
        for i,item in enumerate(row):
            if i >=3:
                item.value = output[a][i]
        a+=1








