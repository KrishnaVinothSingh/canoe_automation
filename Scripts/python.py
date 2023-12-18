# Import CANoe module
import xlwings as xw
from py_canoe import CANoe
from time import sleep as wait

# current excel path##
wb = xw.Book('F:/Python-Vector-CANoe-master/Python_xlwings/xlwings.xlsx')
sheet = wb.sheets['Sheet1']

# create CANoe object
canoe_inst = CANoe()

# create CANoe object
canoe_inst = CANoe()

# open CANoe configuration. Replace canoe_cfg with yours.
canoe_inst.open(canoe_cfg=r'D:/02_CAN_Config/Python/python_access/python_access.cfg')

# print installed CANoe application version
canoe_inst.get_canoe_version_info()

# Start CANoe measurement
canoe_inst.start_measurement()

###  ABS SIGNALS #### 
def set_abs_info(parameter,value):
    
    set_signal=canoe_inst.set_signal_value('CAN', 1, 'ABSInfo',parameter,value)
  
def get_abs_info(parameter):
    
    get_signal=canoe_inst.get_signal_value('CAN', 1, 'ABSInfo', parameter)
    print(get_signal)

###  IC SIGNALS #### 
  
def get_IC_info(parameter):
    
    get_signal=canoe_inst.get_signal_value('CAN', 1, 'IC1_100', parameter)
    print(get_signal)

@xw.func
def hello(name):
    return f'Hello {name}'

###Function call###
set_abs_info(parameter='Front_Speed',value=100)
wait(2)
get_abs_info(parameter='Front_Speed')
set_abs_info(parameter='Front_Speed',value=25)
wait(2)
get_abs_info(parameter='Front_Speed')
wait(2)
get_IC_info(parameter='IndicatedVehSpeed')

#print(type(a))


