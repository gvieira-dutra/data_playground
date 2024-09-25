## Hello Everyone,

This project is about retrieving and manipulating data to extract useful insights for a pharmacogenomics project.

My participation was coding parts of the data manipulation process that would be repetitive and too troublesome to be done by hand so I helped with a few of them.

On the [JSON_Management](https://github.com/gvieira-dutra/data_playground/Pharmacogenomics_data/JSON_Management) folder, you will find the code used to transfer several JSON objects to an excell file. These files can be found [here](https://www.pharmgkb.org/downloads) under the guidelineAnnotations.json.zip downloadable.  
I chose to do this part in C# because althought I've done many software development projects in this language, few of them had much data manipulation. I usually go to Python or R for that purpose so I wanted to change things a bit.  
The goal of this stage was to prepare the data so the Geneticists could analyze them and define next steps based on what was updated since last guideline update.

On the [History_update](https://github.com/gvieira-dutra/data_playground/Pharmacogenomics_data/History_update) folder, you is the part of the project where we needed to retrieve history data about the updates on the guidelines of several medications.  
I created a web scrapper in python that goes to each drug page, checks if there were any updates on 2024, get's the updated information and add it to the excel file.
This will make it easier for the Genticists to analyze these updates and decide if they should modify or keep the current guideline for each drug.
