import os
import re
import win32com.client as w3c
from pathlib import Path
from functions.utils import getChildTasks, getParentTask, getGrandparentTask

def fileExists(path_carpet):
    mppFiles: os._ScandirIterator[str] = []
    masterFile: os.DirEntry[str] = None
    
    with os.scandir(path_carpet) as files:
        for file in files: 
            if file.name.endswith('.mpp'):
                # Example of regular expression : PM-00
                fileName = Path(file.name).stem
                pattern = r'^PM-.*'
                if bool(re.match(pattern, fileName)):
                    mppFiles.append(file)
                else:
                    masterFile = file
    
    return mppFiles, masterFile

def getTaskByName(baseNode: str, file, type):
    mpp: w3c.CDispatch = w3c.Dispatch("MSProject.Application")
    mpp.FileOpen(file.path)
    masterProject = mpp.ActiveProject
      
    tasksMasterProject = masterProject.Tasks  
    taskParent = None    
    listBaseNode = baseNode.split('>')
    
    print('--------------------------------------')
    print(f"VALIDATE FILE {Path(file.name).stem}")
    print('--------------------------------------')
    
    if len(baseNode.strip()) == 0:
        print('Base node is empty')
        print('--------------------------------------')
        return None
  
    indexBaseNode = 0
    for node in listBaseNode:
        task = next((x for x in tasksMasterProject 
                        if x is not None 
                        and x.Name is not None 
                        and x.Name.strip() == node.strip()), 
                    None)
        if task is not None :
            if indexBaseNode == 0 or (getParentTask(task) is not None 
                                        and getParentTask(task).strip() == listBaseNode[indexBaseNode - 1].strip()):
                tasksMasterProject = getChildTasks(task)
                taskParent = task
                print(f"Task: {node.strip()}")
            else:
                print(f"Task: {node.strip()} not found")
                tasksMasterProject = []
                break
        else:
            print(f"Task: {node.strip()} not found")
            tasksMasterProject = []
            break
        indexBaseNode += 1
        
    mpp.FileSave()
    return mpp, taskParent, tasksMasterProject

def updateMasterFile(mppMasterProject, taskParentMasterProject, tasksMasterProject, childs):
    try:
        for child in childs:   
            childTasks: any | list = child['tasks']
   
            tasksByProjectManager = [x for x in tasksMasterProject
                                            if x is not None 
                                            and x.ResourceNames is not None
                                            and x.ResourceNames.strip() == child["fileName"].strip()
                                            and x.OutlineParent is not None
                                            and x.OutlineParent.Name is not None
                                            and x.OutlineParent.Name.strip() == taskParentMasterProject.Name.strip()
                                            ]
            
            print(f"--------------------------------------")
            print(f'{child["fileName"]} : {len(tasksByProjectManager)} tasks by project manager')
            print(f'--------------------------------------')

            # master task is summary
            if tasksByProjectManager is not None and len(tasksByProjectManager) > 0:
                
                for taskProjectManager in tasksByProjectManager:
                    try:  
                        
                        if taskProjectManager is not None or False:
                            childTask = next((x for x in childTasks
                                                if x is not None 
                                                and x.Name is not None 
                                                and x.Name == taskProjectManager.Name 
                                                and getParentTask(x) == getParentTask(taskProjectManager)
                                                and getGrandparentTask(x) == getGrandparentTask(taskProjectManager)
                                                ), 
                                            None)
                            masterTask = next((x for x in tasksMasterProject 
                                                    if x is not None 
                                                    and x.Name is not None
                                                    and x.Name == taskProjectManager.Name 
                                                    and getParentTask(x) == getParentTask(taskProjectManager)
                                                    and getGrandparentTask(x) == getGrandparentTask(taskProjectManager)
                                                    ), 
                                                None)
                            
                            if childTask is not None and masterTask is not None and childTask.PercentComplete is not None:                                                    
                                masterTask.PercentComplete = childTask.PercentComplete
                                print(f"Child task {childTask.Name} ({childTask.PercentComplete}% Complete)")
                            else:
                                print(f"Child task {childTask.Name} not percent complete")  
                            
                            if masterTask.Summary is True:                              
                                groupedChildTasks = getChildTasks(childTask)
                                groupedMasterTasks = getChildTasks(masterTask)
                                
                                for task in groupedMasterTasks: 
                                    relatedTask = next((x for x in groupedChildTasks 
                                                            if x is not None 
                                                            and x.Name is not None 
                                                            and x.Name == task.Name 
                                                            and getParentTask(x) == getParentTask(task)
                                                            and getGrandparentTask(x) == getGrandparentTask(task)
                                                            ), 
                                                        None)

                                    if relatedTask is not None:
                                        if relatedTask.PercentComplete is not None:                                                    
                                            task.PercentComplete = relatedTask.PercentComplete
                                            print(f"        => {relatedTask.Name} ({relatedTask.PercentComplete}% Complete)")
                                        else:
                                            print(f"        => {relatedTask.Name} not percent complete")    
                                    else:
                                        print(f"        => {task.Name} not found")
                                                                                 
                        else:
                            print(f"Task project manager: {taskProjectManager.Name} not found")
                                                                                                        
                    except Exception as e:
                        print(f'Error project manager: {str(e)}') 
                
                print(f"Saving master file")   
                mppMasterProject.FileSave()         
                                 
            else:
                print(f'Task {child["fileName"]} not found') 
                 
    except Exception as e:
        print(f'Error: {str(e)}')
        
    finally: 
        mppMasterProject.FileSave()         
        mppMasterProject.Quit()
        
