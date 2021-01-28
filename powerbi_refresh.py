import time
import os
import sys
import argparse
import psutil
import logging
from pywinauto.application import Application
from pywinauto import timings

PROCNAME = "PBIDesktop.exe"
REFRESH_INTERVAL = 60*2
FORMAT = "%(asctime)-15s %(message)s"
TIMEOUT = 60


def type_keys(string, element):
  """Type a string char by char to Element window"""
  for char in string:
    element.type_keys(char)


def main():
  logging.basicConfig(format=FORMAT, filename="log.log", level=logging.INFO, datefmt="%d-%m-%Y %H:%M:%S")
  parser = argparse.ArgumentParser()
  parser.add_argument("workbook", help = "Path to .pbix file")
  parser.add_argument("excel", help = "Path to .xlsm file")
  args = parser.parse_args()

  timings.after_clickinput_wait = 1  # Default 0.09

  for proc in psutil.process_iter():
    if proc.name() == PROCNAME:
      logging.error("PBI Desktop is already open, please close it.")
      return False

  # Start PBI and open the workbook
  logging.info("")
  logging.info("Power BI Starting...")
  os.system('start "" "' + args.workbook + '"')
  time.sleep(5)

  # Connect pywinauto
  logging.info("Connecting to Power BI...")
  app = Application(backend='uia').connect(path=PROCNAME)
  win = app.window(title_re='.*Power BI Desktop')
  win.wait("ready", timeout=TIMEOUT)
  win.Save.click()  # Save the window
  win.wait("enabled", timeout=TIMEOUT)
  win.set_focus()  # Make window visible
  win.wait("enabled", timeout=TIMEOUT)
  win.maximize()  # Put window in fullscreen 
  win.wait("enabled", timeout=TIMEOUT)
  win.Visualisations.click()  # Close Visualisations toolbox
  win.wait("enabled", timeout=TIMEOUT)
  win.Champs.click()  # Close Fields toolbox
  win.wait("enabled", timeout=TIMEOUT)
  win.Accueil.click_input()  # Check the Home toolbar is selected
  win.wait("enabled", timeout=TIMEOUT)

  # Refresh
  last_modified = os.path.getmtime(args.excel)  # Get excel last modification timestamp
  try:
    while True:
      time.sleep(2)  # To avoid high CPU load
      
      tmp = os.path.getmtime(args.excel)
      if tmp == last_modified:
        continue  # If excel timestamps are the same, the excel file has not been modified. Waiting.
        
      last_modified = tmp
      logging.info("Refreshing")
      try:
        win.set_focus()  # Make window visible
        win.Accueil.click_input()  # Double-check Home toolbar selected
        win.Actualiser.click_input()  # Refresh data
        win.Actualiser.Fermer.click_input()  # If errors in data, a popup window needs to be closed
        win.Save.click()  # Save the window
      except Exception as e:
        logging.error(e)  # We don't really care, the program must not crash
  except KeyboardInterrupt:
    logging.info("Got stop request, stopping now...")
    
  #Close
  logging.info("Exiting")
  win.close()


if __name__ == '__main__':
  try:
    main()
  except (SyntaxError, Exception) as e:
    logging.error(e)
    time.sleep(1000*60)
