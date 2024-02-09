
# SurveyWeightingGUI

## Short Description

SurveyWeightingGUI is a web application built to make survey post-stratification weighting easier and more user-friendly (currently only available in Slovene language).

## Installation

To run app locally, first you need to install [R](https://cran.r-project.org/), then run the following code in R console: 

```
# install missing packages when running app locally from GitHub
required_packages <- c("shiny", "shinyWidgets", "shinyjs", "shinycssloaders", "shinyFeedback", "DT", "rhandsontable", "haven", "labelled", "openxlsx", "anesrake", "survey")

for(p in required_packages){
  if(!require(p, character.only = TRUE)) install.packages(p)
}

# run app
shiny::runGitHub(repo = "SurveyWeightingGUI", username = "lukastrlekar", ref = "main", subdir = "Shiny App")
```

App will also be availabe as R package at a later date.


Copyright (C) 2023-2024 Luka Štrlekar, Vasja Vehovar
