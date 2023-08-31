Agregar tiempo de demora 
- Agregar dond se cambia el nombre del recurso
guia de instalación


# ms-project-helper with Python
Project that allows updating % complete of a main file based on child files

:warning: Avoid leaving open the files that will be used

## Use case
You have two project managers that use the same ms project structure and you need to merge the file to
present to the CEO of the company.

## Requirements

- Python >= 3.7 [installation](https://www.digitalocean.com/community/tutorials/install-python-windows-10)
- Install PIP on Windows [installation](https://www.geeksforgeeks.org/how-to-install-pip-on-windows/) 
[others](https://stackoverflow.com/questions/23708898/pip-is-not-recognized-as-an-internal-or-external-command)

### Libraries

- pywin32 [documentation](https://pypi.org/project/pywin32/)


## Parameters
Shipping parameters in the console.

| Parameter| Description| Example|
| ---| ---| ---|
| --source_folder or -sf| Source folder origin| "C:/docs"|
| --operation or -o| Type of operation| update|
| --base_node or -bn| Base node| "Project > System 1"

Command to list required arguments

```python
py main.py -h
```

## Manual start
```python
pip install -r requirements.txt

py main.py -sf "source folder" -o update -bn "Project > System 1"
```

## Steps
1. Create a folder containing the Ms project files, with the following parameters
  - Base file: File to be updated must not conform to the format of the child file and must be unique
  - Child file: It must contain the name (the format must be "PM-cell name 'resource name (Project manager name)'", example "PM-Pamela Ramirez") of the block of tasks that you want to update

    Project example: [docs.zip](https://github.com/usil/ms-project-helper/files/12458870/docs.zip)

  :warning: File names must not contain commas or special characters.

  :warning: There must be a single base file.

  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/d10986d1-c708-4144-859b-05193cef5e21" width="100%">
  </p>
  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/d7989fb4-748e-4645-9133-cb5019ff219a" width="100%">
  </p>

2. Run command

:warning: Remember to close all files .mpp before execute this script

```python
pip install -r requirements.txt

py main.py -sf "source folder" -o update -bn "Project > System 1"
```

  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/98f90075-4b55-4b61-ae82-4fd1cb5ac75e" width="100%">
  </p>

  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/7f3a7fe7-1a0f-438d-a0a1-65938730f4b7" width="100%">
  </p>

3. Result 
  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/05ee0f7b-0a40-43c0-a0df-8d3314547d26" width="100%">
  </p>


## Note (developers)
- Lines 76 and 77, there is the field names of the resources of the file 'functions/mpp.py'.


## Contributors

<table>
  <tbody>
    <td>
      <img src="https://avatars.githubusercontent.com/u/77288944?v=4" width="100px;"/>
      <br />
      <label><a href="https://github.com/madeliyricra">Madeliy Ricra</a></label>
      <br />
    </td>  
  </tbody>
</table>

## Documentation

- https://www.geeksforgeeks.org/how-to-install-pip-on-windows/
- https://j2logo.com/python/listar-directorio-en-python/
- https://docs.python.org/3/library/re.html
- https://rctorr.wordpress.com/2014/01/16/procesando-parametros-en-la-linea-de-comando-en-python/
