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


def type_keys(string, element):
  """Type a string char by char to Element window"""
  for char in string:
    element.type_keys(char)


def main():
  logging.basicConfig(format=FORMAT, filename="log.log", level=logging.INFO, datefmt="%d-%m-%Y %H:%M:%S")
  parser = argparse.ArgumentParser()
  parser.add_argument("workbook", help = "Path to .pbix file")
  args = parser.parse_args()

  timings.after_clickinput_wait = 1

  for proc in psutil.process_iter():
    if proc.name() == PROCNAME:
      logging.error("PBI Desktop is already open, please close it.")
      return False

  # Start PBI and open the workbook
  logging.info("\n")
  logging.info("Starting Power BI")
  os.system('start "" "' + args.workbook + '"')
  time.sleep(5)

  # Connect pywinauto
  logging.info("Identifying Power BI window")
  app = Application(backend='uia').connect(path=PROCNAME)
  win = app.window(title_re='.*Power BI Desktop')
  win.wait("ready", timeout=30)
  win.Save.click()  # Save the window
  win.wait("enabled", timeout=30)
  win.set_focus()  # Make window visible
  win.wait("enabled", timeout=30)
  win.maximize()  # Put window in fullscreen 
  win.wait("enabled", timeout=30)
  win.Visualisations.click()  # Close Visualisations toolbox
  win.wait("enabled", timeout=30)
  win.Champs.click()  # Close Fields toolbox
  win.wait("enabled", timeout=30)
  win.Accueil.click_input()  # Check the Home toolbar is selected
  win.wait("enabled", timeout=30)

  # Refresh
  try:
    while True:
      logging.info("Refreshing")
      win.set_focus()  # Make window visible
      win.Accueil.click_input()  # Double-check Home toolbar selected
      win.Actualiser.click_input()  # Refresh data
      try:
        win.Actualiser.Fermer.click_input()  # If errors in data, a popup window needs to be closed
      except:
        pass
      win.Save.click()  # Save the window
      time.sleep(REFRESH_INTERVAL)
  except KeyboardInterrupt:
    pass
    

  #Close
  logging.info("Exiting")
  win.close()

  # Force close
  #for proc in psutil.process_iter():
  #  if proc.name() == PROCNAME:
  #    proc.kill()

    
if __name__ == '__main__':
  try:
    main()
  except (SyntaxError, Exception) as e:
    logging.error(e)
    time.sleep(1000*60)