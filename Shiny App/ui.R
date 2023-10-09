suppressPackageStartupMessages({
  library(shiny)
  library(shinyWidgets)
  library(shinyjs)
  library(shinycssloaders)
  library(shinyFeedback)
  library(DT)
  library(rhandsontable)
})

options(spinner.type = 8)

shinyUI(
  navbarPage(title = "SurveyWeightingGUI", id = "nav_menu", collapsible = TRUE,
             shinyjs::useShinyjs(),
             includeCSS("www/style.css"),
             tags$head(tags$script("
                        Shiny.addCustomMessageHandler('resetValue', function(variableName) {
                        Shiny.onInputChange(variableName, null);
                        });
                      ")),
             
             # Home tab / instructions -------------------------------------------------
             tabPanel(title = "Domov", icon = shiny::icon("home"),
                      fluidRow(
                        shinyFeedback::useShinyFeedback(), # include shinyFeedback
                        
                        style = "background-color: #2e6da4;
                                  color: white;
                                  min-height: 20px;
                                  padding: 19px;
                                  margin-bottom: 20px;
                                  border-radius: 4px;
                                  margin-left: 16px;
                                  margin-right: 16px;",
                        fluidRow(
                          column(4),
                          column(4,
                                 br(),
                                 p("SurveyWeightingGUI",
                                   style = "font-size: 36px;
                                            font-weight: bold;
                                            text-align: center;"),
                                 br(),
                                 p("Aplikacija za enostavno in hitro uteževanje anketnih podatkov",
                                   style = "font-size: 24px;
                                            font-weight: normal;
                                            text-align: center;"),
                                 br(),
                                 br()),
                          column(4)),
                        fluidRow(
                          column(3),
                          column(3,
                                 actionButton("read_more_weighting",
                                              label = "Preberi več o uteževanju",
                                              width = "100%",
                                              style = "background-color: #2e6da4;
                                                       border-color: white;
                                                       color: white;
                                                       font-size: 18px;")),
                          column(3,
                                 actionButton("start_using",
                                              label = "Začni z uporabo",
                                              width = "100%",
                                              style = "background-color: white;
                                                       border-color: white;
                                                       color: #2e6da4;
                                                       font-size: 18px;"),
                                 br(),
                                 br(),
                                 br(),
                                 br()),
                          column(3))),
                      
                      fluidRow(
                        column(4, offset = 4, align = "center",
                               h3(strong("3 koraki uteževanja z aplikacijo")),
                               hr())),
                      fluidRow(
                        column(12, align = "center",
                               h4(HTML("<center><b>1.</b></br>Naloži podatke</center>")),
                               br(),
                               h4(HTML("<center><b>2.</b></br>Izberi uteževalne spremenljivke in vnesi populacijske (ciljne) deleže</center>"))),
                        br(),
                        column(4, offset = 4,
                               h4(HTML("<center><b>3.</b></br>Izvedi uteževanje</center>")),
                               p("Po izvedbi uteževanja prenesete datoteko z utežmi in jih uporabite v analizi podatkov v svojem najljubšem programu za statistično analizo.
                                    Dodatno lahko že v sami aplikaciji preverite spremembo povprečij in deležev spremenljivk po uteževanju.",
                                 class = "banner-text"),
                        )),
                      fluidRow(align = "center",
                               br(),
                               hr(),
                               p(HTML('&copy; 2023 Luka Štrlekar, Vasja Vehovar &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                       <a href="mailto:ls9889@student.uni-lj.si">Kontakt <i class="fa fa-envelope"></i></a> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                       <a href="https://github.com/lukastrlekar/SurveyWeightingGUI" target="_blank" rel="noopener noreferrer">
                                              Izvorna koda <i class="fa fa-github"></i></a>')))),
  
             # Data input tab -------------------------------------------------------
             tabPanel(title = "Nalaganje podatkov", value = "upload_data_tab", icon = shiny::icon("upload"),
                      column(4,
                             wellPanel(
                               h4(strong("Nalaganje podatkov")),
                               br(),
                               radioButtons("raw_data_file_type", 
                                            label = NULL, 
                                            choices = c("SPSS (.sav)" = "sav",
                                                        "CSV" = "csv",
                                                        "TXT" = "txt",
                                                        "R (.rds)" = "rds"), 
                                            inline = TRUE),
                               conditionalPanel(condition = 'input.raw_data_file_type == "csv" | input.raw_data_file_type == "txt"',
                                                hr(),
                                                column(6,
                                                       selectInput("decimal_separator", 
                                                                   label = HTML("<small>Izberi decimalno ločilo:</small>"),
                                                                   choices = c("Vejica (,)" = ",",
                                                                               "Pika (.)" = "."))),
                                                column(6,
                                                       selectInput("delimiter", 
                                                                   label = HTML("<small>Izberi delimiter:</small>"),
                                                                   choices = c("Podpičje (;)" = ";",
                                                                               "Vejica (,)" = ",",
                                                                               "Tabulator (tab)" = "\t",
                                                                               "Presledek ( )" = " ",
                                                                               "Pipa (|)" = "|"))),
                                                column(12,
                                                       textInput("na_strings",
                                                                 label = HTML("<small>Obravnavaj kot manjkajoče vrednosti:</small>"),
                                                                 value = "-1, -2, -3, -4, -5, -96, -97, -98, -99")),
                                                column(12,
                                                       hr())),
                               br(),
                               fileInput("upload_raw_data",
                                         label = NULL,
                                         buttonLabel = "Naloži podatke",
                                         accept = c(".sav", ".csv", ".txt", ".rds")),
                               p(HTML("<small><b>Navodila:</b> Podatki morajo biti v tabelarični obliki, kjer je vsaka enota v svoji vrstici in vsaka
                                              spremenljivka v svojem stolpcu. Največja dovoljena velikost naloženih podatkov je <b>100 MB</b>.</br> </br>
                                              <b>Priporočila:</b> Pred uporabo aplikacije v <i>Predogled naloženih podatkov</i> preverite, če
                                              so bili podatki prebrani v pravilni obliki.</small>"))
                             ),
                             actionButton("jump_to_select_variables_tab", label = "Naprej >", class = "btn-next")),
                      column(8,
                             h4(strong("Predogled naloženih podatkov")),
                             br(),
                             
                             hidden(p("Podatkov ni bilo mogoče naložiti. Preverite, da ste izbrali pravo končnico datoteke ali pravi separator oziroma decimalno ločilo.",
                                      id = "message_wrong_raw_data",
                                      style = "color:red;")),
                             
                             shinycssloaders::withSpinner(
                               DT::DTOutput("raw_data_table")))),
             
             # Select weighting variables tab ------------------------------------------------------
             tabPanel(title = "Izbira spremenljivk", value = "select_variables_tab", icon = shiny::icon("hand-pointer"),
                      column(12, h4(strong("Izbira spremenljivk za uteževanje")),
                             br()),
                      column(9, 
                             wellPanel(fluidRow(
                               column(6,
                                      pickerInput(
                                        inputId = "one_dim_variables",
                                        label = HTML("Izberi spremenljivke: </br>
                                                       <small>(enodimenzionalna kontingenčna tabela)</small>"),
                                        choices = NULL,
                                        multiple = TRUE,
                                        options = pickerOptions(actionsBox = TRUE,
                                                                liveSearch = TRUE))),
                               column(6,
                                      pickerInput(
                                        inputId = "two_dim_variables",
                                        label = HTML("Izberi pare spremenljivk: </br>
                                                       <small>(dvodimenzionalna kontingenčna tabela)</small>"),
                                        choices = NULL,
                                        options = pickerOptions(actionsBox = TRUE,
                                                                liveSearch = TRUE,
                                                                maxOptions = 2),
                                        multiple = TRUE),
                                      
                                      disabled(actionButton("add", label = NULL, icon = shiny::icon("plus"))))),
                               fluidRow(
                                 column(12,
                                        htmlOutput("selected_variables")),
                                 column(12,
                                        br(),
                                        p(HTML("<small><b>Navodila in priporočila:</b> Izberite spremenljivke in morebitne kombinacije (dvo-dimenzionalne interakcije) spremenljivk, s katerimi boste izvedli uteževanje. To so običajno opisne (nominalne)
                                                        socio-demografske spremenljivke, ki nimajo velikega deleža manjkajočih vrednosti in za
                                                        katere so na voljo zanesljivi in posodobljeni podatki o sestavi (deleži kategorij) ciljne populacije 
                                                        (recimo iz uradnih statističnih virov). Primeri takih spremenljivk so: spol, starost, regija in kraj bivanja, izobrazba, zaposlitveni status ipd.
                                                        Pred uteževanjem morate za te izbrane spremenljivke za vsako kategorijo poiskati ciljne populacijske deleže.</small>"))))),
                             fluidRow(column(12,
                                             checkboxInput("display_tables_preview",
                                                           label = "Prikaži predogled tabel", 
                                                           value = FALSE)),
                                      wellPanel(style = "background: #FFFFFF; border-color: #FFFFFF; box-shadow: none",
                                                conditionalPanel(condition = "input.display_tables_preview == 1",
                                                                 div(class = "second-tab",
                                                                     shinycssloaders::withSpinner(
                                                                       uiOutput("preview_tables"))))))),
                      column(3, 
                             wellPanel(
                               p(strong("Nastavitve prikaza tabel")),
                               checkboxInput("drop_zero_categories",
                                             label = "Ne prikazuj praznih vsebinskih kategorij", value = TRUE),
                               checkboxInput("drop_missings_categories",
                                             label = "Ne prikazuj praznih manjkajočih kategorij", value = TRUE)),
                             actionButton("jump_to_input_margins_tab", label = "Naprej >", class = "btn-next"))),
             
             # Input margins tab ------------------------------------------------------
             tabPanel(title = "Vnos margin", value = "input_margins_tab", icon = shiny::icon("pen"),
                      column(12, h4(strong("Vnos populacijskih (ciljnih) margin")),
                             br()),
                      column(10,
                             wellPanel(
                               p(strong("Izberite mesto vnosa populacijskih (ciljnih) margin:")),
                               radioGroupButtons("select_which_input",
                                                 label = NULL,
                                                 choices = c("&nbsp Vnos margin v aplikaciji &nbsp <i class='fa fa-arrow-pointer'></i>" = "app_input",
                                                             "&nbsp Vnos margin v ločeni Excel datoteki &nbsp <i class='fa fa-file-excel'></i>" = "excel_input"),
                                                 justified = TRUE,
                                                 individual = TRUE,
                                                 selected = character(0),
                                                 checkIcon = list(
                                                   yes = tags$i(class = "fa-regular fa-circle-dot", 
                                                                style = "color: black"),
                                                   no = tags$i(class = "fa fa-circle-o", 
                                                               style = "color: black"))))),
                      
                      conditionalPanel(condition = 'input.select_which_input == "excel_input"',
                                       column(10, align="center",
                                              wellPanel(
                                                p(strong("Vnos populacijskih (ciljnih) margin v ločeni Excel datoteki"), style = "font-size:16px;"),
                                                br(),
                                                p(HTML("<b>Navodila za vnos:</b> S klikom na spodnji gumb prenesite Excel datoteko, v katero nato za vse izbrane spremenljivke
                                                        v stolpec <i>Vnos populacijskih margin (%)</i> vnesite populacijske (ciljne) margine v odstotkih. V stolpcu <i>Populacijski (ciljni) %</i>
                                                        se samodejno preračunajo deleži, ki bodo uporabljeni v uteževanju (glede na morebitne manjkajoče vrednosti in za
                                                        zagotovilo, da bo vsota vedno 100%). Ko končate z vnosom, Excel datoteko shranite in jo naložite v aplikacijo v zavihku <i>Uteževanje</i>.")),
                                                p(HTML('<small>Opombe: Manjkajoče vrednosti (<i>Missing</i>) se obravnavajo samodejno, glede na postopek, ki ga uporablja
                                                       <a href="https://casovnice.cdi.si/uploadi/editor/doc/1492779173DOC1_R7_weights.pdf" target="_blank">Evropska družboslovna raziskava</a>.
                                                       Morebitna opozorila in priporočila o združevanju kategorij so povzeta iz splošnih pripročil za tovrstno uteževanje 
                                                       (glej npr. <a href="https://www.surveypractice.org/article/2953-practical-considerations-in-raking-survey-data" target="_blank">Battaglia in drugi (2009)</a>,
                                                       <a href="https://surveyinsights.org/wp-content/uploads/2014/07/Full-anesrake-paper.pdf" target="_blank">Pasek in drugi (2014)</a>,
                                                       <a href="https://www150.statcan.gc.ca/n1/en/pub/12-001-x/2010002/article/11376-eng.pdf?st=ZBMtNMXC" target="_blank">Särndal in Lundström (2010)</a>).
                                                       Odsotnost majhnih kategorij je zaželena saj pospeši konvergenco, manjša možnost nastanka zelo neenakomernih (variabilnih) oz. nestabilnih uteži – 
                                                       visoko variabilne uteži povečujejo vzorčno varianco in lahko povzročijo bolj nestabilne točkovne ocene.</small>')),
                                                br(),
                                                disabled(
                                                  downloadButton("download_excel_input", 
                                                                 label = HTML("&nbspPrenesi datoteko za vnos margin"), 
                                                                 icon = shiny::icon("file-excel fa-solid", style = "color: white;"),
                                                                 class = "btn-primary")))),
                                       column(2, 
                                              actionButton("jump_to_weighting_tab", label = "Naprej >", class = "btn-next", style = "position: fixed"))),
                      
                      conditionalPanel(condition = 'input.select_which_input == "app_input"',
                                       column(10, align = "center",
                                              wellPanel(
                                                p(strong("Vnos populacijskih (ciljnih) margin v aplikaciji"), style = "font-size:16px;"),
                                                br(),
                                                p(HTML("<b>Navodila za vnos:</b> V modro označen prostor v spodnjih tabelah vnesite populacijske (ciljne) margine za vsako kategorijo spremenljivke v odstotkih (%).
                                                      V stolpcu <i>Populacijski (ciljni) %</i>
                                                      se samodejno preračunajo deleži, ki bodo uporabljeni v uteževanju (glede na morebitne manjkajoče vrednosti in za
                                                      zagotovilo, da bo vsota vedno 100%). Ko končate z vnosom, lahko izvedete uteževanje v zavihku <i>Uteževanje</i>.")),
                                                p(HTML('<small>Opombe: Manjkajoče vrednosti (<i>Missing</i>) se obravnavajo samodejno, glede na postopek, ki ga uporablja
                                                       <a href="https://casovnice.cdi.si/uploadi/editor/doc/1492779173DOC1_R7_weights.pdf" target="_blank">Evropska družboslovna raziskava</a>.
                                                       Morebitna opozorila in priporočila o združevanju kategorij so povzeta iz splošnih pripročil za tovrstno uteževanje 
                                                       (glej npr. <a href="https://www.surveypractice.org/article/2953-practical-considerations-in-raking-survey-data" target="_blank">Battaglia in drugi (2009)</a>,
                                                       <a href="https://surveyinsights.org/wp-content/uploads/2014/07/Full-anesrake-paper.pdf" target="_blank">Pasek in drugi (2014)</a>,
                                                       <a href="https://www150.statcan.gc.ca/n1/en/pub/12-001-x/2010002/article/11376-eng.pdf?st=ZBMtNMXC" target="_blank">Särndal in Lundström (2010)</a>).
                                                       Odsotnost majhnih kategorij je zaželena saj pospeši konvergenco, manjša možnost nastanka zelo neenakomernih (variabilnih) oz. nestabilnih uteži – 
                                                       visoko variabilne uteži povečujejo vzorčno varianco in lahko povzročijo bolj nestabilne točkovne ocene.</small>'))),
                                              fluidRow(align = "left",
                                                       actionButton("render_input_tables",
                                                                    label = "Ponastavi vnos v izbrani tabeli",
                                                                    class = "btn-secondary"))),
                                       column(2, 
                                              wellPanel(
                                                checkboxInput("download_inputed_tables_checkbox",
                                                              label = HTML("Prenesi datoteko z vnešenimi marginami </br>
                                                                          <small>(shranite datoteko z vnešenimi marginami za kasnejšo uporabo)</small>")),
                                                conditionalPanel(condition = "input.download_inputed_tables_checkbox == 1",
                                                                 downloadButton("download_inputed_tables", 
                                                                                label = HTML("&nbspPrenesi datoteko"), 
                                                                                icon = shiny::icon("file-code"),
                                                                                class = "btn-secondary"),
                                                                 helpText(HTML("<small>Prenesla se bo datoteka s končnico .Rdata, ki jo ročno ne boste mogli urejati.</small>")))),
                                              actionButton("jump_to_weighting_tab2", label = "Naprej >", class = "btn-next")),
                                       
                                       column(12,
                                              fluidRow(
                                                wellPanel(style = "background: #FFFFFF; border-color: #FFFFFF; box-shadow: none",
                                                          div(class = "second-tab",
                                                              uiOutput("hot_tables"))))))),
             
             # Weighting tab ---------------------------------------------------------------
             tabPanel(title = "Uteževanje", value = "weighting_tab", icon = shiny::icon("scale-unbalanced"),
                      column(12, h4(strong("Izračun poststratifikacijskih uteži z metodo raking")),
                             br()),
                      column(4,
                             wellPanel(
                               uiOutput("which_weighting_vars"),
                               
                               hidden(p(HTML("Naložite Excel (.xlsx) ali .Rdata datoteko, ki ste jo prenesli v zavihku <i>Vnos populacijskih margin</i>."),
                                        id = "message_wrong_file",
                                        style = "color:red;")),
                               hidden(p(HTML("Naložite podatke v zavihku <i>Nalaganje podatkov</i>."),
                                        id = "message_no_file",
                                        style = "color:red;")),
                               hidden(p("Oblika datoteke se ne ujema s predvideno. Naložite pravo Excel datoteko.",
                                        id = "message_wrong_file_structure",
                                        style = "color:red;")),
                               hidden(p("Spremenljivke v podani datoteki ne obstajajo v naloženih podatkih.",
                                        id = "message_no_variables_present",
                                        style = "color:red;")),
                               hidden(p("Določene spremenljivke v podani datoteki ne obstajajo v naloženih podatkih.
                                        Te spremenljivke so bile izključene iz uteževanja.",
                                        id = "message_some_variables_not_present",
                                        style = "color:red;"))),
                             
                             wellPanel(
                               pickerInput(
                                 inputId = "weights_variables",
                                 label = "Izberi spremenljivke za uteževanje",
                                 choices = NULL,
                                 multiple = TRUE,
                                 options = list(`actions-box` = TRUE))),
                             
                             wellPanel(
                               p(strong("Nastavitve uteževanja")),
                               radioButtons("package",
                                            label = HTML('<span style="font-weight:normal">Izberi paket, s katerim se izvede uteževanje:</span>'),
                                            choices = list("anesrake" = "anesrake",
                                                           "survey" = "survey"),
                                            inline = TRUE),
                               fluidRow(
                                 column(6,
                                        checkboxInput("case_id_selection",
                                                      label = HTML("Izberi enolični identifikator enot (<i>case ID</i>)")),
                                       conditionalPanel(condition = "input.case_id_selection == 1",
                                                        pickerInput(
                                                          inputId = "case_id_variable",
                                                          label = HTML("<small>Izberi spremenljivko, ki predstavlja enolični identifikator enot:</small>"),
                                                          choices = NULL,
                                                          options = pickerOptions(
                                                            liveSearch = TRUE,
                                                            maxOptions = 1),
                                                          multiple = TRUE))),
                                column(6,
                                       checkboxInput("base_weights_selection",
                                                     label = "Upoštevaj tudi privzete uteži"),
                                       conditionalPanel(condition = "input.base_weights_selection == 1",
                                                        p(HTML("<small>Izbirni vnos, če so na voljo vzorčne uteži, stratifikacijski popravek ali druga informacija o verjetnosti vzorčenja, ki jo je treba upoštevati pred izvedbo uteževanja.</small>")),
                                                        pickerInput(inputId = "base_weights_variable",
                                                                    label = HTML("<small>Izberi spremenljivko, ki predstavlja uteži:</small>"),
                                                                    choices = NULL,
                                                                    options = pickerOptions(
                                                                      liveSearch = TRUE,
                                                                      maxOptions = 1),
                                                                    multiple = TRUE)))),
                              
                               checkboxInput("cut_weights_iterative",
                                             label = "Reži uteži pri vsaki iteraciji procedure"),
                               conditionalPanel(condition = "input.cut_weights_iterative == 1",
                                                fluidRow(
                                                  column(4,
                                                         numericInput("cap_weights",
                                                                      label = HTML("<small>Maksimalna vrednost uteži:</small>"),
                                                                      value = 5, min = 1, step = 0.5)),
                                                  column(8,
                                                         p(HTML(
                                                           "<small>Priporočena možnost, saj je z vidika optimizacije variance optimalna.
                                                            V redkih primerih lahko privede do neveljavnih (skoraj ničelnih) uteži (<0.01).
                                                            Takrat se priporoča povečanje max dovoljene vrednosti uteži ali rezanje uteži na koncu procedure raking.</small>"))))),
                               
                               checkboxInput("cut_weights_after",
                                             label = "Uteži poreži na koncu procedure"),
                               conditionalPanel(condition = "input.cut_weights_after == 1",
                                                fluidRow(
                                                  column(6,
                                                         numericInput("min_weight",
                                                                      label = HTML("<small>Minimalna utež:</small>"),
                                                                      value = 0.1, min = 0.001, step = 0.1)),
                                                  column(6,
                                                         numericInput("max_weight",
                                                                      label = HTML("<small>Maksimalna utež:</small>"),
                                                                      value = 5, min = 1, step = 0.5)),
                                                  column(12,
                                                         p(HTML("<small>Zaradi postopka normalizacije lahko vrednosti porezanih uteži rahlo
                                                                       presežejo nastavljeno min in max vrednost.</small>"))))),
                               checkboxInput("convergence_input",
                                             label = "Nastavi kriterij za konvergenco"),
                               conditionalPanel(condition = "input.convergence_input == 1",
                                                conditionalPanel(condition = 'input.package == "anesrake"',
                                                                 p(HTML("<small>Raking algoritem skonvergira, ko zadnja ponovitev predstavlja manj kot nastavljeno odstotno izboljšanje glede na prejšnjo iteracijo.</small>")),
                                                                 column(4,
                                                                        numericInput("convergence_criterion",
                                                                                     label = NULL,
                                                                                     value = 0.01, min = 0.00001, step = 0.001))),
                                                conditionalPanel(condition = 'input.package == "survey"',
                                                                 p(HTML("<small>Raking algoritem skonvergira, če je največja sprememba v vnosu v tabeli manjša od nastavljene vrednosti. Če je nastavljena vrednost < 1, se obravnava kot del skupne vzorčne uteži.</small>")),
                                                                 column(8,
                                                                        numericInput("convergence_criterion_epsilon",
                                                                                     label = NULL,
                                                                                     value = 1e-10, min = 1e-10, step = 0.001)))),
                               br(),
                               actionButton("run_weighting",
                                            label = HTML("&nbspZaženi uteževanje"),
                                            class = "btn-primary",
                                            icon = shiny::icon("play"),
                                            width = "100%"))),
                      
                      column(8,
                             div(class = "first-tab",
                                 tabsetPanel(id = "weighting_panel",
                                             tabPanel("Povzetek uteževanja", icon = shiny::icon("circle-info"),
                                                      br(),
                                                      
                                                      hidden(p(HTML("Vklopili ste možnost <i>Upoštevaj tudi privzete uteži</i>, vendar niste izbrali spremenljivke uteži.
                                                          Privzete uteži se zato niso upoštevale. <br/>"),
                                                               id = "message_no_base_weights_selected",
                                                               style = "color:red;")),
                                                      
                                                      # htmlOutput("weighting_variables_input_messages"),
                                                      #tags$head(tags$style("#code_summary{overflow-y:scroll; max-height: 200px; background: white; border: white}")),
                                                      shinycssloaders::withSpinner(
                                                        htmlOutput("code_summary")),
                                                      br(),
                                                      fluidRow(
                                                        column(4,
                                                               uiOutput("jump_to_download_weights_ui")),
                                                        column(4,
                                                               uiOutput("weighting_diagnostic_download_ui"))),
                                                      br(),
                                                      p(strong("Frekvence uteževalnih spremenljivk"), style = "font-size:16px;"),
                                                      div(class = "second-tab",
                                                          uiOutput("weighting_variables_tables"))),
                                             
                                             tabPanel("Porazdelitev uteži", icon = shiny::icon("chart-column"),
                                                      br(),
                                                      tableOutput("summary_stat"),
                                                      plotOutput("weights_plot", width = "60%")),
                                             
                                             tabPanel("Predogled uteži", icon = shiny::icon("table-columns"),
                                                      br(),
                                                      DT::DTOutput("weights_table", width = "40%")),
                                             
                                             tabPanel("Prenos uteži",
                                                      value = "download_weights_tab",
                                                      icon = shiny::icon("file-arrow-down"),
                                                      
                                                      br(),
                                                      p(strong("Prenos datoteke z utežmi")),
                                                      radioButtons("select_weights_file_type",
                                                                   label = NULL,
                                                                   choices = c("TXT" = "txt",
                                                                               "CSV" = "csv",
                                                                               "SPSS (.sav)" = "sav"),
                                                                   inline = TRUE),
                                                      conditionalPanel(condition = 'input.select_weights_file_type == "csv" | input.select_weights_file_type == "txt"',
                                                                       hr(),
                                                                       fluidRow(
                                                                         column(3,
                                                                                selectInput("decimal_separator_weights", 
                                                                                            label = HTML("<small>Izberi decimalno ločilo:</small>"),
                                                                                            choices = c("Pika (.)" = ".",
                                                                                                        "Vejica (,)" = ","))),
                                                                         column(3,
                                                                                selectInput("delimiter_weights", 
                                                                                            label = HTML("<small>Izberi delimiter:</small>"),
                                                                                            choices = c("Podpičje (;)" = ";",
                                                                                                        "Vejica (,)" = ",",
                                                                                                        "Tabulator (tab)" = "\t",
                                                                                                        "Presledek ( )" = " ",
                                                                                                        "Pipa (|)" = "|")))),
                                                                       fluidRow(
                                                                         column(12,
                                                                                checkboxInput("quote_col_names_weights",
                                                                                              label = 'Imena stolpcev naj bodo v dvojnih narekovajih ("")',
                                                                                              width = "100%")))),
                                                      fluidRow(
                                                        column(6,
                                                               br(),
                                                               downloadButton("download_weights",
                                                                              label = HTML("&nbspPrenesi datoteko uteži"),
                                                                              icon = shiny::icon("file-arrow-down"),
                                                                              class = "btn-primary",
                                                                              style = "width: 100%;"))),
                                                      
                                                      conditionalPanel(condition = "input.package == 'survey'",
                                                                       hr(),
                                                                       fluidRow(
                                                                         column(6,
                                                                                p(strong(HTML("Prenos <tt>svydesign</tt> R objekta iz paketa <i>survey</i>"))),
                                                                                p(HTML("<small>Prenesite <tt>svydesign</tt> objekt za nadaljne analize v R s paketom survey. V objektu so shranjene informacije
                                                               o uteževanju, ki omogočajo pravilen (natančen) izračun standardnih napak, ne zgolj aproksimativen (npr. s Taylorjevo linearizacijo).</small>")),
                                                                                br(),
                                                                                downloadButton("download_survey_design",
                                                                                               label = HTML("&nbspPrenesi svydesign objekt"),
                                                                                               icon = shiny::icon("file-code"),
                                                                                               class = "btn-secondary",
                                                                                               style = "width: 100%;"))))
                                                      
                                             ))))),
             
             # Analyses tab ------------------------------------------------------------
             tabPanel(title = "Analize", icon = shiny::icon("calculator"),
                      column(12, h4(strong("Osnovna primerjava spremenljivk pred in po uteževanju")),
                             br()),
                      column(12,
                             wellPanel(
                               fluidRow(
                                 column(5,
                                        radioButtons("select_which_weights",
                                                     label = "Katere poststratifikacijske uteži želite upoštevati pri izračunu?",
                                                     choices = c("Novo izračunane (v zavihku Uteževanje)" = "new_weights",
                                                                 "Obstoječe, v naloženih podatkih" = "included_weights")),
                                        conditionalPanel(condition = "input.select_which_weights == 'included_weights'",
                                                         pickerInput(
                                                           inputId = "analyses_weight_variable",
                                                           label = HTML("<small>Izberi spremenljivko poststratifikacijskih uteži:</small>"),
                                                           choices = NULL,
                                                           options = pickerOptions(
                                                             liveSearch = TRUE,
                                                                    maxOptions = 1),
                                                                  multiple = TRUE,
                                                                  width = "60%")),
                                        htmlOutput("weights_message"),
                                        br(),
                                        checkboxInput("adjust_p_values",
                                                      label = HTML("Popravi p vrednosti zaradi večkratnega preizkušanja domnev")),
                                        conditionalPanel(condition = "input.adjust_p_values == 1",
                                                         selectInput("select_p_adjust_method",
                                                                     label = "Izberi metodo popravljanja p vrednosti:",
                                                                     choices = list("Holm" = "holm",
                                                                                    "Hochberg" = "hochberg",
                                                                                    "Hommel" = "hommel",
                                                                                    "Bonferroni" = "bonferroni",
                                                                                    "Benjamini & Hochberg (BH/FDR)" = "BH",
                                                                                    "Benjamini & Yekutieli (BY)" = "BY"),
                                                                     selected = 1, multiple = FALSE, width = "50%"))
                                        ),
                                 column(7,
                                        disabled(radioButtons("se_calculation",
                                                              label = "Izberi način izračuna standardne napake (SE):",
                                                              choices = c("Aproksimativen izračun SE s Taylorjevo linearizacijo za razmernostno cenilko" = "taylor_se",
                                                                          "Natančen izračun SE s paketom survey" = "survey_se"))),
                                        p(HTML('<small><b>Preverjanje statistične značilnosti vpliva uteževanja:</b> Za številske spremenljivke se preverja ničelna domneva, da je uteženo povprečje enako neuteženemu. Uporabljen je t-test
                                        za en vzorec.
                                        Za vsako kategorijo spremenljivke v frekvenčnih tabelah se preverja ničelno domnevo, da je utežen delež enak neuteženemu. Uporabljen je 
                                        z-test za delež, kjer se standardna napaka prav tako oceni po isti formuli.</small>')))
                               ))),
                      
                      column(12,
                             div(class = "first-tab",
                             tabsetPanel(tabPanel(title = "Izračun opisnih statistik",
                                                  br(),
                                                  column(5,
                                                         pickerInput(
                                                           inputId = "numeric_variables",
                                                           label = HTML('Izberi številske spremenljivke:<br/><span style="font-weight:normal"><small>(intervalna ali razmernostna merska lestvica)</small></span>'),
                                                           choices = NULL,
                                                           multiple = TRUE,
                                                           options = pickerOptions(actionsBox = TRUE,
                                                                                   liveSearch = TRUE),
                                                           width = "100%")),
                                                  column(1,
                                                         div(style = "top: 45px; position:relative;",
                                                             actionButton("run_numeric_variables",
                                                                          label = "Izračunaj",
                                                                          class = "btn-primary"))),
                                                  column(6,
                                                         div(style = "top: 45px; position:relative;",
                                                             uiOutput("download_numeric_analyses_ui"))),
                                                  column(12,
                                                         htmlOutput("message_numeric"),
                                                         shinycssloaders::withSpinner(
                                                           uiOutput("analyses_numeric_tables")))),
                                         
                                         tabPanel(title = "Izračun frekvenc",
                                                  br(),
                                                  column(5,
                                                         pickerInput(
                                                           inputId = "factor_variables",
                                                           label = HTML('Izberi spremenljivke:<br/><span style="font-weight:normal"><small>(nominalna, ordinalna, intervalna ali razmernostna merska lestvica)</small></span>'),
                                                           choices = NULL,
                                                           multiple = TRUE,
                                                           options = pickerOptions(actionsBox = TRUE,
                                                                                   liveSearch = TRUE),
                                                           width = "100%")),
                                                  column(1,
                                                         div(style = "top: 45px; position:relative;",
                                                             actionButton("run_factor_variables",
                                                                          label = "Izračunaj",
                                                                          class = "btn-primary"))),
                                                  column(6,
                                                         div(style = "top: 45px; position:relative;",
                                                             uiOutput("download_factor_analyses_ui"))),
                                                  column(12,
                                                         htmlOutput("message_factor"),
                                                         shinycssloaders::withSpinner(
                                                           uiOutput("analyses_factor_tables"))))
                             ))))))