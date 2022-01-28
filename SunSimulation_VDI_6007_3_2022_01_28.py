# -*- coding: utf-8 -*-
# SunSimulation_VDI_6007_3.py
# Version 2022_01_28

# To calculate sun position and solar radiation to surfaces for a year
# According too VDI 6007 - 3
# Based on Global Horizontal Irradiance Clear Sky Models

# Designed by Engin Bagda
# see: explanations in www.heizen-co2-sparen.de

def SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg, Normal_direct, Horizontal_global, Horizontal_sky,SSW):

# Incidence Angle IA , angle between direction of sun radiation and surface normal in Radiant

    IA_rad = math.cos(Zenith_rad) * math.cos(Surface_Tilt_rad)
    IA_rad = IA_rad + math.sin(Surface_Tilt_rad) * math.sin(Zenith_rad) * math.cos(abs(Azimuth_rad - Surface_Azimuth_rad))

    if IA_rad > 1: IA_rad = 1
    if IA_rad < 0: IA_rad = 0

    IA_rad = math.acos(IA_rad)  # because IA is cos(IA) = cos(zenith)*cos (Surf TA) + ....

    R_iso = 0.182 * ( 1.178 * (1+math.cos(Surface_Tilt_rad))+(math.pi - Surface_Tilt_rad)*math.cos(Surface_Tilt_rad)+ math.sin(Surface_Tilt_rad))

    if Elevation_deg > 21.5:

        R_180 = -21* ( 1-4*(21.5)/90)
    else:
        R_180 = -21 * (1 - 4 * (Elevation_deg) / 90)

    R_WBL_0 = 6 * (1 - ((Elevation_deg-15)/15)**2)

    if Elevation_deg > 30: R_WBL_0 = 0

    R_WBL_1 = -6.5*(1-((math.degrees(Surface_Tilt_rad)-40)/45)**2)

    if R_WBL_1 > 0: R_WBL_1 = 0

    R_WBL =(-64.5*math.sqrt(abs(math.sin(math.radians(Elevation_deg)))) + R_WBL_0) * (1 - math.degrees(Surface_Tilt_rad)/180) + R_WBL_1

    R_WBNL = 13 * (1 - math.sin(2*Surface_Tilt_rad))

    R_IA = (126.5 - 60*math.sin(math.radians(Elevation_deg))) * (((math.cos(IA_rad)+0.7)/1.7)**2)

    R_sky = R_iso + ((R_180 + R_WBL + R_WBNL + R_IA)/100)  #  *math.sin(Surface_Tilt_rad)

# Direct sun irradiation to surface
    Surface_direct = (Horizontal_direct / math.cos(Zenith_rad)) * math.cos(IA_rad) * SSW

# Sky irradiance
    Surface_sky = Horizontal_sky * R_sky*SSW + Horizontal_sky * R_iso * (1-SSW)

# Reflection from ground
    Reflection_ground = 0.2
    Surface_reflec = Horizontal_global * 0.5 * Reflection_ground * (1 - math.cos(Surface_Tilt_rad))*SSW

# Diffuse irradiance
    Surface_diffuse = Surface_reflec + Surface_sky

# Total irradiance on surface
    Surface_total = Surface_direct + Surface_diffuse

    return { 'Surface_total': Surface_total,
             'R_sky': R_sky,
             'Surface_diffuse' : Surface_diffuse  }

# Main run

import math
from typing import Any, Union

import xlsxwriter # to export to excel

Surface_Azimuth_deg_variable = 180  # 90° east, 180° south, 270 ° West, 0 North
Surface_Tilt_deg_variable = 30

Latitude_deg = 49.0562   # For Mannheim
Longitude_deg = 8.5585 # for Mannheim to calculate in true local time

Height = 98  # Altitude of the place above sea level in m for Mannheim

TF_min = 3.7    # According to VDI 6007-3 table 1
TF_max = 6.1    # According to VDI 6007-3 table 1

TF_ave = (TF_min + TF_max)/2
TF_ampl = (TF_max - TF_min)/2

Surface_Azimuth_rad_variable = math.radians(Surface_Azimuth_deg_variable)
Surface_Tilt_rad_variable = math.radians(Surface_Tilt_deg_variable) # Vertical surface: pi/2, Horizontal surface: zero
Latitude_rad = math.radians(Latitude_deg)

print()
print ("   Surface azimuth : %5.0f"% (math.degrees(Surface_Azimuth_rad_variable)))
print ("Surface tilt angle : %5.0f"% (math.degrees(Surface_Tilt_rad_variable)))
Time_zone = float(input("         Time_zone :     "))  # to calculate in true local time set timezone = 0,  for Germany at summertime 2

print("            TF_min : %5.1f"% (TF_min))
print("            TF_max : %5.1f"% (TF_max))
SSW =  float(input("      Selected SSW :     "))
print()

day_number = int(input("  Selected day Nr. :   "))

workbook = xlsxwriter.Workbook('SunRadiation.xlsx') # Excel output file name

worksheet_RAD = workbook.add_worksheet('Solar_radiance')
worksheet_Daily = workbook.add_worksheet('Daily-Values') # Incidance angle

# Values for irradiation in W/m2
worksheet_RAD.write(0, 0, "True local")
worksheet_RAD.write(1, 0, "hour")
worksheet_RAD.write(0, 1, "True local")
worksheet_RAD.write(1, 1, "minute")

worksheet_RAD.write(0, 2, "Local")
worksheet_RAD.write(1, 2, "hour")
worksheet_RAD.write(0, 3, "Local")
worksheet_RAD.write(1, 3, "minute")

worksheet_RAD.write(0, 4, "Irrad.")
worksheet_RAD.write(0, 5, "Irrad.")
worksheet_RAD.write(0, 6, "Irrad.")
worksheet_RAD.write(0, 7, "Irrad.")
worksheet_RAD.write(0, 8, "Irrad.")
worksheet_RAD.write(0, 9, "Irrad.")
worksheet_RAD.write(0, 10, "Irrad.")
worksheet_RAD.write(0, 11, "Irrad.")

worksheet_RAD.write(1, 4, "Horiz.glob")
worksheet_RAD.write(1, 5, "Horiz.sky")
worksheet_RAD.write(1, 6, "east")
worksheet_RAD.write(1, 7, "south")
worksheet_RAD.write(1, 8, "west")
worksheet_RAD.write(1, 9, "north")
worksheet_RAD.write(1, 10, "surface.total")
worksheet_RAD.write(1, 11, "surface.diff")

worksheet_Daily.write(0, 1, "Irradiation daily in kWh/m2/dayr")
worksheet_Daily.write(1, 0, "Day Number")
worksheet_Daily.write(1, 1, " TF")
worksheet_Daily.write(1, 3, "Horiz.glob")
worksheet_Daily.write(1, 4, "Horiz.sky")
worksheet_Daily.write(1, 5, "east")
worksheet_Daily.write(1, 6, "south")
worksheet_Daily.write(1, 7, "west")
worksheet_Daily.write(1, 8, "north")
worksheet_Daily.write(1, 9, "surface.total")
worksheet_Daily.write(1, 10, "surface.diff")

Time_step = 3  # for Worksheet lines in 5 Minute steps, starting at excel line 3
Day_step  = 1  # for Worksheet lines in daily steps, starting at excel line 3

for day in range (1,366,1):

    Horizontal_global_sum = Horizontal_sky_sum = Surface_diffuse_sum = 0
    Surface_total_sum = East_total_sum = South_total_sum = West_total_sum = North_total_sum = 0

    TF = TF_ave - TF_ampl*math.cos((math.pi * 2 / 365) * day)

    if day == day_number:
        print("     Calculated TF : %5.1f" % (TF))
        TF =  float(input("    TF for the day :   "))

        print()
        print("   Local time  True time     Zenith   Azimuth   Horizontal Horizontal    Surface   Surface    Elevation   Azimuth     R")
        print("                                                  Global       Sky        Total    Diffuse               difference    ")
        print("    hh:mm       hh:mm         Deg      Deg         W/m2       W/m2         W/m2      W/m2       Deg         Deg    ")

    # Solar constant in W/m2 Median 1367.7 W/m2, according to Reno, Hansen, Stein equation 11
    SC = 1367.7 * (1 + 0.033 * math.cos(math.pi * 2 * day / 365))

    for hour in range (4, 21, 1): # in local true time respective solar time

        for Minute in range(0, 60, 5):

            # according to Reno, Hansen, Stein equation 6 to slice the year in radians/day, calculation start at 00:00
            fy_rad = (2 * math.pi / 365) * (day - 1)

            # Declination
            Declination_rad = 0.006918 - 0.399912 * math.cos(fy_rad) + 0.070257 * math.sin(fy_rad) - 0.006758 * math.cos(
                2 * fy_rad) + 0.000907 * math.sin(2 * fy_rad) - 0.002697 * math.cos(3 * fy_rad) + 0.00148 * math.sin(
                3 * fy_rad)

            # Hour_angle
            # earth rotation 15 degrees per hour

            Hour_Angle_deg = ((hour*60+Minute)/60 - 12) * 15
            Hour_Angle_rad = math.radians(Hour_Angle_deg)

            # Zenith Angle of sun, at zenith ZA=0, at horizon ZA = 90°

            Cos_Zenith_rad = math.sin(Latitude_rad) * math.sin(Declination_rad) + math.cos(Latitude_rad) * math.cos(
                Declination_rad) * math.cos(Hour_Angle_rad)

            Zenith_rad = math.acos(Cos_Zenith_rad)

            Elevation_deg = 90 - math.degrees(Zenith_rad)

            # Azimut angle of sun

            Cos_Azimuth_rad = ((math.cos(Zenith_rad) * math.sin(Latitude_rad))
                               - math.sin(Declination_rad)) / (math.cos(Latitude_rad) * math.sin(Zenith_rad))

            if Hour_Angle_rad == 0:
                Azimuth_deg = 180

            if Hour_Angle_rad > 0:
                Azimuth_deg = 180 + math.degrees(math.acos(Cos_Azimuth_rad))

            if Hour_Angle_rad < 0:
                Azimuth_deg = 180 - math.degrees(math.acos(Cos_Azimuth_rad))

            Azimuth_rad = math.radians(Azimuth_deg)

            # Air mass

            Air_mass = 1 / (0.9 + 9.4 * math.sin(math.radians(Elevation_deg)))

            # Absorption by transmission trough atmosphere

            if math.degrees(Zenith_rad) < 90:  # after Zenith angle=90 is sun set
                TaM = 1.294 + 2.4417 * 10 ** -2 * Elevation_deg - 3.973 * 10 ** -4 * Elevation_deg ** 2
                TaM = TaM + 3.8034 * 10 ** -6 * Elevation_deg ** 3 - 2.2145 * 10 ** -8 * Elevation_deg ** 4
                TaM = (TaM + 5.8832 * 10 ** -11 * Elevation_deg ** 5) * (0.506 - 1.0788 * 10 ** -2 * TF)

                # Direct normal irradiance (Bestrahlungsstaerke auf Flächennormale, Sonne)

                Normal_direct = SC * math.exp(-TF * Air_mass * math.exp(-Height / 8000))

                # Direct horizontal irradiance (Horizontale Bestrahlungsstärke, Sonne)

                Horizontal_direct = Normal_direct * math.cos(Zenith_rad)

                # Horizontal irradiance from sky (Horizontale Bestrahlungsstaerke, Himmel)

                Horizontal_sky = 0.5 * SC * (math.cos(Zenith_rad)) * (
                            TaM - math.exp(-TF * Air_mass * math.exp(-Height / 8000)))

                # Global horizontal irradiance (Horizontale Globale Bestrahlungsstaerke)

                Horizontal_global = Horizontal_direct + Horizontal_sky

                Horizontal_global_sum += Horizontal_global
                Horizontal_sky_sum += Horizontal_sky

                # East
                Surface_Azimuth_rad = math.pi / 2  # 90° oriented to east
                Surface_Tilt_rad = math.pi/2 # Verticular surface

                Surface_irrad_result = SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg,
                                                    Normal_direct, Horizontal_global, Horizontal_sky,SSW)

                if day == day_number:
                    worksheet_RAD.write(Time_step, 6, int(Surface_irrad_result['Surface_total']))

                East_total_sum += Surface_irrad_result['Surface_total']

                # South
                Surface_Azimuth_rad = math.pi  # 180° oriented to south
                Surface_Tilt_rad = math.pi / 2  # Verticular surface

                Surface_irrad_result = SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg,
                                                    Normal_direct,
                                                    Horizontal_global, Horizontal_sky,SSW)
                if day == day_number:
                    worksheet_RAD.write(Time_step, 7, int(Surface_irrad_result['Surface_total']))

                South_total_sum += Surface_irrad_result['Surface_total']

                # West
                Surface_Azimuth_rad = math.pi * 3 / 2  # 270° oriented to west
                Surface_Tilt_rad = math.pi / 2  # Verticular surface

                Surface_irrad_result = SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg,
                                                    Normal_direct,
                                                    Horizontal_global, Horizontal_sky,SSW)
                if day == day_number:
                    worksheet_RAD.write(Time_step, 8, int(Surface_irrad_result['Surface_total']))

                West_total_sum += Surface_irrad_result['Surface_total']
    
                # North

                Surface_Azimuth_rad = math.pi * 4 / 2  # 360° oriented to north
                Surface_Tilt_rad = math.pi / 2  # Verticular surface

                Surface_irrad_result = SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg,
                                                    Normal_direct,
                                                    Horizontal_global, Horizontal_sky,SSW)
                if day == day_number:
                    worksheet_RAD.write(Time_step, 9, int(Surface_irrad_result['Surface_total']))

                North_total_sum += Surface_irrad_result['Surface_total']

                # Surface variable

                Surface_Azimuth_rad = Surface_Azimuth_rad_variable
                Surface_Tilt_rad = Surface_Tilt_rad_variable

                Surface_irrad_result = SurfaceIrrad(Surface_Azimuth_rad, Zenith_rad, Azimuth_rad, Elevation_deg,
                                                    Normal_direct, Horizontal_global, Horizontal_sky,SSW)

                if day == day_number:
                    worksheet_RAD.write(Time_step, 10, int(Surface_irrad_result['Surface_total']))
                    worksheet_RAD.write(Time_step, 11, int(Surface_irrad_result['Surface_diffuse']))

                Surface_total_sum += Surface_irrad_result['Surface_total']
                Surface_diffuse_sum += Surface_irrad_result['Surface_diffuse']

                # Calculation Local time
                # Equation of Time (EoT)


                x = (2 * math.pi / 365) * ((day - 81))
                EoT = 9.87 * math.sin(2 * x) - 7.53 * math.cos(x) - 1.5 * math.sin(x)

                Time_offset = ((Time_zone * 15 - Longitude_deg) / 15) * 60 + EoT  # in Minutes

                Local_time = (hour * 60 + Time_offset) / 60

                Local_Hour = int(Local_time)
                Local_Minute = (Local_time - Local_Hour) * 60 + Minute

                if Local_Minute > 59.5:
                    Local_Hour = hour + 2
                    Local_Minute = Local_Minute - 60
                if day == day_number:
                    if hour == 12 and Minute == 0: print()
                    if Minute == 0:
                        print(
                            "%6.0f %2.0f %8.0f %2.0f %12.1f  %8.1f  %8.0f   %8.0f     %8.0f  %8.0f   %8.0f    %8.0f %8.2f" %
                            (
                            Local_Hour, Local_Minute, hour, Minute, math.degrees(Zenith_rad), math.degrees(Azimuth_rad),
                            Horizontal_global, Horizontal_sky, Surface_irrad_result['Surface_total'],
                            Surface_irrad_result['Surface_diffuse'],
                            Elevation_deg, abs(math.degrees((Azimuth_rad - Surface_Azimuth_rad))),
                            Surface_irrad_result['R_sky']))

                    if hour == 12 and Minute == 0: print()

                if day == day_number:

                    if Minute == 0:
                        worksheet_RAD.write(Time_step, 0, hour)

                    worksheet_RAD.write(Time_step, 1, Minute)

                    if int(Local_Minute) < 5:
                        worksheet_RAD.write(Time_step, 2, Local_Hour)

                    worksheet_RAD.write(Time_step, 3, int(Local_Minute))
                    worksheet_RAD.write(Time_step, 4, int(Horizontal_global))
                    worksheet_RAD.write(Time_step, 5, int(Horizontal_sky))

                    Time_step = Time_step + 1

    if day == day_number:
        worksheet_RAD.write(Time_step + 2, 0, "Total")

        worksheet_RAD.write(Time_step + 2, 1, "in kWh/m2/day")
        worksheet_RAD.write(Time_step + 2, 4, (Horizontal_global_sum) / 12 / 1000)  # because calculation every 5 minute
        worksheet_RAD.write(Time_step + 2, 5, (Horizontal_sky_sum) / 12 / 1000)     # Sum is divided by 12 for per hour
        worksheet_RAD.write(Time_step + 2, 6, (East_total_sum) / 12 / 1000)
        worksheet_RAD.write(Time_step + 2, 7, (South_total_sum) / 12 / 1000)
        worksheet_RAD.write(Time_step + 2, 8, (West_total_sum) / 12 / 1000)
        worksheet_RAD.write(Time_step + 2, 9, (North_total_sum) / 12 / 1000)
        worksheet_RAD.write(Time_step + 2, 10, (Surface_total_sum) / 12 / 1000)
        worksheet_RAD.write(Time_step + 2, 11, int(Surface_diffuse_sum) / 12 / 1000)

        print()
        print(" Irradiation per day in kWh/m2/day            %8.2f   %8.2f     %8.2f  %8.2f" % (Horizontal_global_sum/1000/12, Horizontal_sky_sum/1000/12, Surface_total_sum/1000/12, Surface_diffuse_sum/1000/12))

    worksheet_Daily.write(Day_step + 2, 0, (day))
    worksheet_Daily.write(Day_step + 2, 1, (TF))
    worksheet_Daily.write(Day_step + 2, 3, (Horizontal_global_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 4, (Horizontal_sky_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 5, (East_total_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 6, (South_total_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 7, (West_total_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 8, (North_total_sum) / 12 / 1000)

    worksheet_Daily.write(Day_step + 2, 9, (Surface_total_sum) / 12 / 1000)
    worksheet_Daily.write(Day_step + 2, 10, (Surface_diffuse_sum) / 12 / 1000)

    Day_step += 1

workbook.close()