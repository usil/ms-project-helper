import win32com.client as w3c

def createMpp(path):
  try:
    mpp = w3c.Dispatch("MSProject.Application")
    try:
      mpp.FileNew()
      proj = mpp.ActiveProject
      task = proj.Tasks.Add("Tarea 1")
      task.Duration = "5"
      task.Start = "1/1/2020"
      task.Finish = "5/1/2020"
      task.ActualStart = "1/1/2020"
      task.ActualFinish = "5/1/2020"
      task.PercentComplete = 100
      task.PercentWorkComplete = 100
      task.PercentPhysicalWorkComplete = 100

      print (proj.Tasks.Count)
    except Exception as e:
      print(str(e))
    mpp.FileSaveAs(path)
    mpp.Quit()
  except Exception as  e:
      print(str(e))

def readMpp(path):
  try:
    mpp = w3c.Dispatch("MSProject.Application")
    try:
      mpp.FileOpen(path)
      proj = mpp.ActiveProject
      print (proj.Tasks.Count)
      for task in proj.Tasks:
        print (task.Name)
    except Exception as e:
      print(str(e))
  except Exception as  e:
      print(str(e))


def updateMpp(path):
  try:
    mpp = w3c.Dispatch("MSProject.Application")
    try:
      mpp.FileOpen(path)
      proj = mpp.ActiveProject

      proj.Tasks.Add("Tarea 4")

      task = proj.Tasks.Add("Tarea 2")
      task.Duration = "5"
      task.Start = "1/1/2020"
      task.Finish = "5/1/2020"
      task.ActualStart = "1/1/2020"
      task.ActualFinish = "5/1/2020"
      task.PercentComplete = 100
      task.PercentWorkComplete = 100

      print (proj.Tasks.Count)
    except Exception as e:
      print(str(e))
    mpp.FileSave()
    mpp.Quit()
  except Exception as  e:
      print(str(e))