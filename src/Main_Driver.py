# Main Driver for Optimization Software
#
# Assumptions
# 1. Delta V loss is 20% of ideal delta V
# 2. The balloon is not rotating WRT the Earth
# 3.

# Good Combination
# First Stage
#   Delta V = 43%
#   Isp = 250 s
#   Propellant Mass Fraction = .85
# Second Stage
#   Delta V = 57%
#   Isp = 270 s
#   Propellant Mass Fraction = .9
# Glow ~= 110 kg

from GLOW_Optimization_Driver import GLOW_Optimization_Driver
import xlsxwriter
from tqdm import tqdm

# Initializes number of stages
Num_Stages = 2

# Velocity split range and the step size
VelocitySplitRange = [40, 60]
VelocitySplitStep = 1

# Isp range and step size
# Applied to all stages
IspRange = [240, 280]
IspStep = 10

# Propellant mass fraction range and step size
# Applied to all stages
PropMassFractRange = [85, 90]
PropMassFractStep = 5

# Orbit Altitude
Altitude = 130000

# Payload Mass
PayloadMass = .2

# Maximum GLOW (Kg)
MaxGLOW = 500

# Minimum Glow
MinGLOW = 100

print("Simulation Started")
Data = GLOW_Optimization_Driver(Num_Stages, VelocitySplitRange, VelocitySplitStep,
                                IspRange, IspStep, PropMassFractRange, PropMassFractStep,
                                Altitude, PayloadMass, MaxGLOW, MinGLOW)

# For printing data from simulation to excel sheet
workbook = xlsxwriter.Workbook('Data_Stages_' + str(Num_Stages) + '.xlsx')
worksheet = workbook.add_worksheet("Sheet 1")

print("Saving Data")

# Excel sheet column titles
title = ["Velocity Split of Stage 1",
            "Velocity Split of Stage 2",
            "Isp of Stage 1",
            "Isp of Stage 2",
            "Propellant Mass Fraction of Stage 1",
            "Propellant Mass Fraction of Stage 2",
            "GLOW"]
worksheet.write_row(0, 0, title)

# prints data to excel sheet
sheet = 1
col = 0
for row in tqdm(range(len(Data))):
    worksheet.write_row((row+sheet) % 1048576, col, Data[row])

    if (((row+sheet+1) % 1048576) == 0) & (row > 2):
        sheet = sheet + 1
        worksheet = workbook.add_worksheet("Sheet " + str(sheet))
        worksheet.write_row(0, 0, title)
workbook.close()
print("Simulation complete")
