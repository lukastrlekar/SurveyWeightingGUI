# temporary check to install packages when running app locally form GitHub

required_packages <- c("shiny", "shinyWidgets", "shinyjs", "shinycssloaders", "DT", "rhandsontable", "haven", "labelled", "openxlsx", "anesrake")

for(p in required_packages){
  if(!require(p, character.only = TRUE)) install.packages(p)
}
