from ctypes import Structure, windll, c_uint, sizeof, byref
import threading
import schedule
import requests
import json
from time import sleep


# Load client info from .env file
client = {}
try:
    with open(".env", 'r') as f:
        client = json.load(f)
except FileNotFoundError:
    raise Exception(".env File not found")

# print settings
print("OpenHab item: %s" % client['OHItem'])
print("Polling every %d seconds" % client['pollSec'])
print("Inactivity threshold: %d seconds" % client['inactiveThreshold'])
print("Forced server update every %d seconds" % client['ServerUpdateRate'])
print("-------------------------------------")


# Windows last input class
class INACTIVITY:
  activeStatus = False                      # Inactive status flag. Computer is currently inactive if True

  #Print out every n seconds the idle time, when moving mouse, this should be < 10
  # https://stackoverflow.com/a/29730972
  def status(self):
    # threading.Timer(client['pollSec'], self.calculate).start()
    iTime = self.get_idle_duration()
    activeStatus = False
    if iTime < client['inactiveThreshold']:
      activeStatus = True
    #print ("Active: %d  Time: %d" %(activeStatus, iTime))
    return activeStatus

  # Data structure for Window's Last User input info
  class LASTINPUTINFO(Structure):
      _fields_ = [
          ('cbSize', c_uint),
          ('dwTime', c_uint),
      ]

  # Get idle duration from Windows7
  def get_idle_duration(self):
      lastInputInfo = self.LASTINPUTINFO()
      lastInputInfo.cbSize = sizeof(lastInputInfo)
      windll.user32.GetLastInputInfo(byref(lastInputInfo))
      millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
      return millis / 1000.0


inac = INACTIVITY()
lastStatus = False


# Send activity status to server
def sendIdleStatus(x):
  print("Sending %r to server" % x)
  if x == True: # if
    data = 'ON'
  elif x == False:
    data = 'OFF'
  else:
    raise Exception("Invalid data to sendIdleStatus")

  # Send activity to server
  try:
    myresponce = requests.post(client['RestURL'] + 'items/' + client['OHItem'], data, auth=(client['User'],client['Secret']), timeout=3.0)
  except (requests.ConnectTimeout, requests.ConnectionError) as e:
    print ("Connection error")
    print(str(e))
  except (requests.ReadTimeout, requests.Timeout) as e:
    print ("Request Timedout")
    print(str(e))
  except requests.RequestException as e:
    print("Request: General Error")
    print(str(e))
  else:
    print(myresponce.text)


# Test for status change and send status update if needed
def activeLogicEdge():
    global lastStatus
    Status = inac.status()
    if (Status != lastStatus):
        sendIdleStatus(Status)
    lastStatus = Status

# Always sends status update.
def activeLogicPeriodic():
    global lastStatus
    Status = inac.status()
    sendIdleStatus(Status)
    lastStatus = Status



# Schedule tasks
if __name__ == "__main__":
    activeLogicEdge()
    schedule.every(client['pollSec']).seconds.do(activeLogicEdge)
    schedule.every(client['ServerUpdateRate']).seconds.do(activeLogicPeriodic)


# Defind action when this program closes
import atexit
def closing():
    print("Sending last update and closing")
    sendIdleStatus(False)   # Send False update to make sure server status doesn't stay ON permanently

atexit.register(closing)    # register closing fuction




# Main loop
while 1:
  sleep(1)
  schedule.run_pending()
