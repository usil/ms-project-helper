# ms-project-helper with Python

## Requirements

- Python >= 3.7
- Install PIP on Windows [installation](https://www.geeksforgeeks.org/how-to-install-pip-on-windows/)
- Create a folder containing the Ms project files, with the following parameters
  1. Base file: File to be updated must not conform to the format of the child file and must be unique
  2. Child file: It must contain the name (the format must be "Number. Name") of the block of tasks that you want to update

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

## Manual start (developers)
```python
pip install -r requirements.txt

py main.py -sf "source folder" -o update
```

## Manual start (production)
```python
```

## Example 

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