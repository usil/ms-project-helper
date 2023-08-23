def getChildTasks(task):
    childTasks = []
    for subTask in task.OutlineChildren:
        childTasks.append(subTask)
        childTasks.extend(getChildTasks(subTask))
    return childTasks

def getGrandparentTask(task):
    if task.OutlineLevel > 2: # Grandparent task
        parentTask = task.OutlineParent
        grandparentTask = parentTask.OutlineParent
        return grandparentTask.Name
    else:
        return None

def getParentTask(task):
    if task.OutlineLevel > 1: # Parent task
        parentTask = task.OutlineParent
        return parentTask.Name
    else:
        return None
    
def getFieldsTask(task):
    for attr in dir(task):
        if not attr.startswith("__") and not callable(getattr(task, attr)):
            print(attr)   
        break     