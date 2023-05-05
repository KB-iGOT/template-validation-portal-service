# template-validation-portal-service

Repository for backend service of Data upload and Validation tool

Code pushes to be done in the `dev` branch only.


## Requirements
1. Python dependencies
2. MongoDB data restore

## Python dependencies

There are two ways to install python dependencies :-


1. Conda and environment.yml file (recommended):-

```
conda env create -f environment.yml
conda activate templateValidation
```

Note :- Please refer to below link for installing conda in ubuntu

https://docs.conda.io/projects/conda/en/latest/user-guide/install/linux.html


2. Virtual env and requirement.txt file :-
```
python -m venv env_name
source env_name/bin/activate
pip install -r requirements.txt
```

## MongoDB data restore

Use following command to restore mongoDB dump :-

```
cd data
mongorestore --host localhost --port 27017 --db templateValidation --gzip ./
```

## Execution 
```
cd apiServices/src/main/
python app.py
```
