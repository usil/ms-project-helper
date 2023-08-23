import argparse
import os
from functions.mpp import fileExists, updateMasterFile

def main(arguments):
    source_folder: str = arguments.source_folder
    operation: str = arguments.operation
       
    files, masterFile = fileExists(source_folder)
    print(f'Master file: {masterFile.path}')
    
    if operation == "update": 
        updateMasterFile(masterFile, files)
    else:
        print("Operation not supported")

if __name__ == '__main__':
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
    
    args: argparse.Namespace = parser.parse_args()
    
    if args.source_folder and os.path.exists(args.source_folder):
        main(args)
    else:
        print('Error: The source folder does not exist')
    