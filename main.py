import argparse
import os
from pathlib import Path
from functions.mpp import fileExists, updateMasterFile, getTaskByName

def main(arguments):
    sourceFolder: str = arguments.source_folder
    operation: str = arguments.operation
    baseNode: str = arguments.base_node
    
    files, masterFile = fileExists(sourceFolder)

    print(f'MASTER FILE: {masterFile.name}')
    print(f'BASE NODE: {baseNode}')
    tasksChilds = []
    
    for file in files:
        mpp, taskParent, tasks = getTaskByName(baseNode, file, '')
        if taskParent is not None:
            fileName = Path(file.name).stem
            # Project manager fileName
            fileName = fileName.replace('PM-', '')
            tasksChilds.append({
                'file': file,
                'fileName': fileName,
                'taskParent': taskParent,
                'tasks': tasks
            })
        else:
            print(f'File {file.path} not found')
                       
    mppMasterProject, taskParentMasterProject, tasksMasterProject = getTaskByName(baseNode, masterFile, 'master')
    
    print('--------------------------------------')
    print(f"UPDATE MASTER FILE {Path(masterFile.name).stem}")
    if operation == "update": 
        updateMasterFile(mppMasterProject, taskParentMasterProject, tasksMasterProject, tasksChilds)
    else:
        print("Operation not supported")

if __name__ == '__main__':
    print(f'Remember to close all files .mpp before execute this script')
    parser: argparse.ArgumentParser = argparse.ArgumentParser()
    parser.add_argument('-sf', 
                        '--source_folder', 
                        required=True,
                        help='Source folder origin',
                        type=str)
    
    parser.add_argument('-o',
                        '--operation',
                        required=True,
                        help='Type of operation',
                        choices=["update"])
    
    parser.add_argument('-bn',
                        '--base_node',
                        required=True,
                        help='Base node',
                        type=str)
    
    args: argparse.Namespace = parser.parse_args()
    
    if args.source_folder and os.path.exists(args.source_folder):
        main(args)
    else:
        print('Error: The source folder does not exist')
    