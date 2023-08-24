# ms-project-helper with Python
Project that allows updating % complete of a main file based on child files

:warning: Avoid leaving open the files that will be used

## Use case
You have two project managers that use the same ms project structure and you need to merge the file to
present to the CEO of the company.

## Requirements

- Python >= 3.7
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

Command to list required arguments

```python
py main.py -h
```

## Steps
1. Create a folder containing the Ms project files, with the following parameters
  - Base file: File to be updated must not conform to the format of the child file and must be unique
  - Child file: It must contain the name (the format must be "Number. Name") of the block of tasks that you want to update

  :warning: There must be a single base file

  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/a1ae5410-a657-4a95-8a5e-03624c180f88" width="100%">
  </p>
  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/f0d050ab-3117-4a47-a820-3560ac848e4f" width="100%">
  </p>

2. Run command
  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/1b0aa3b4-96c3-4ca6-bb8d-03ff94f3fec7" width="100%">
  </p>
3. Result 
  <p align="center">
    <img src="https://github.com/usil/ms-project-helper/assets/77288944/9e04c557-35ab-47fc-a2b4-989de093bcc2" width="100%">
  </p>


## Manual start (developers)
```python
pip install -r requirements.txt

py main.py -sf "source folder" -o update
```

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
