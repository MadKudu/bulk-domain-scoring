# Bulk domain scoring

This little scripts is meant to help batching huge xlsx files to score each
lines.

###### Installation

- Make sure you have python3 installed (within a pyenv/venv is ideal)
- `pip3 install -r requirements.txt`

###### Run
- `filename` the file from which to read from either xls or xlsx
- `score_type` defines wich way to score the records, either email or domain
- `column_idx` is the letter(s) corresponding to the column where the script should retrieve the email or domain information for each row (based on score type)
- `api_key` is the api key of the tenant

`python3 bulk_score.py --filename="file_to_batch.xlsx" --score_type="email" --column_idx="A" --api_key="tenant_api_key"`

###### Run against custom files
- `filename` the file from which to read from either xls or xlsx
- `score_type` defines wich way to score the records, either email or domain
- `column_idx` is the letter(s) corresponding to the column where the script should retrieve the email or domain information for each row (based on score type)
- `api_key` is the api key of the tenant

`python3 bulk_score.py --filename="file_to_batch.xlsx" --score_type="email" --column_idx="A" --api_key="tenant_api_key"`

###### Error handling

This script is state-full and creates a csv file containing all the `results`
as soon as received. So that, in case of failure and once restarted, it will
read the results file content to resume operations.

The corollary is that this script state is reset by removing the
corresponding file in the `results` folder.

###### Zip results back to xlsx
`python3 zip_result --filename="file_to_batch.xlsx"`


###### Limitations
This project was made for a custom file and does not currently support
alternative formats or column order. Feel free to improve! 
