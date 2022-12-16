import os, os.path, sys

from pyaxidraw import axidraw   
ad = axidraw.AxiDraw()


def GoPlot():
    port = sys.argv[1]
    print("Printing to port "+ port)

    file_to_print = sys.argv[2]
    file_to_print = file_to_print.replace("-", " - ")
    file_to_print = file_to_print.replace("!", " ")
    file_to_print = file_to_print.replace("_", "(")
    file_to_print = file_to_print.replace("=", ")")
    print("The path to plot is " +file_to_print)

    ad.plot_setup(file_to_print)
    #set standard plotter options
    ad.options.speed_pendown = 85
    ad.options.speed_penup = 85
    ad.options.accel = 90
    ad.options.pen_pos_down = 25
    ad.options.pen_pos_up = 50
    ad.options.pen_rate_lower = 90
    ad.options.pen_rate_raise = 90
    ad.options.pen_delay_down = 2
    ad.options.auto_rotate = False
    ad.options.reordering = 2
    ad.options.report_time = True
    #ad.options.random_start = True

    ad.options.port = port

    ad.plot_run()


GoPlot()

#Assign port
#You can specify the machine using the USB port enumeration (e.g., 
# COM6 on Windows or /dev/cu.usbmodem1441 on a Mac) or 
# by using an assigned USB nickname
#ad.options.port = "COM5"

#ad.plot_setup("Templates/Casey - Pulsz - Envelope Template1.svg")
#ad.plot_run()