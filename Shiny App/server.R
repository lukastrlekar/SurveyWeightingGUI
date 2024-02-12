suppressPackageStartupMessages({
  library(shiny)
  library(shinyWidgets)
  library(shinyjs)
  library(shinycssloaders)
  library(shinyFeedback)
  library(DT)
  library(rhandsontable)
  
  library(haven)
  library(labelled)
  library(openxlsx)
  library(anesrake)
  library(survey)
})

source("Excel_output_files.R")
source("Data_manipulation.R")
source("Raking_script.R")
source("Analyses_functions.R")

server <- function(input, output, session){
  
  ## Global settings
  options(shiny.sanitize.errors = FALSE,
          shiny.maxRequestSize = 100 * 1024^2) # Change maximum file size upload to 100 MB
          # htmlwidgets.TOJSON_ARGS = list(na = 'string')) # Display missing values as NA instead of blank space (https://stackoverflow.com/questions/58526047/customizing-how-datatables-displays-missing-values-in-shiny)
  
  # Change style of file upload buttons to primary buttons
  runjs("$('#upload_raw_data').parent().removeClass('btn-default').addClass('btn-primary');")
  
  # Next button events
  observeEvent(input$start_using, {
    updateNavbarPage(session = session, inputId = "nav_menu", selected = "upload_data_tab")
  })
  
  observeEvent(input$jump_to_select_variables_tab, {
    updateNavbarPage(session = session, inputId = "nav_menu", selected = "select_variables_tab")
  })
  
  observeEvent(input$jump_to_input_margins_tab, {
    updateNavbarPage(session = session, inputId = "nav_menu", selected = "input_margins_tab")
  })
  
  observeEvent(input$jump_to_weighting_tab, {
    updateNavbarPage(session = session, inputId = "nav_menu", selected = "weighting_tab")
  })
  
  observeEvent(input$jump_to_weighting_tab2, {
    # first check if sum of user inputed values is every table is > 0
    # if not display a message
    results <- unlist(lapply(seq_along(inputed_tables()), function(i) {
      t <- hot_to_r(input[[paste0("input_table_", clean_names()[[i]])]])
      t[nrow(t),ncol(t)] > 0
    }))
    
    if(all(results) == FALSE || length(clean_names()) != length(results)){
      showModal(modalDialog(HTML("<big><strong><center>Vnesite margine v vse tabele!</center></strong></big>"),
                            easyClose = TRUE,
                            footer = modalButton(icon("xmark"))))
      return()
      
    } else {
      updateNavbarPage(session = session, inputId = "nav_menu", selected = "weighting_tab")
    }
  })
  
  # Show more about weighting text
  observeEvent(input$read_more_weighting, {
    showModal(modalDialog(
      title = NULL,
      HTML(
        "<h4><strong><center>Poststratifikacijsko uteževanje anket</center></strong></h4>
          </br>
          S poststratifikacijskim uteževanjem porazdelitve izbranih, običajno demografskih, spremenljivk z vzorca uskladimo z znano populacijsko (ciljno) porazdelitvijo teh spremenljivk.
          V aplikaciji je implementirana v praksi najpogosteje uporabljena metoda uteževanja raking, ki deluje tako, da ponavlja postopek 
          poststratifikacije iterativno za vsako izbrano spremenljivko posebej, dokler se vzorčni deleži ne približajo populacijskim dovolj natančno.
          </br></br>           
          Tako uteževanje vrne uteži, ki zagotovijo, da se uteženi rezultati anket ujemajo s populacijskimi (ciljnimi) porazdelitvami izbranih spremenljivk. Uporaba teh uteži v statističnih analizah
          tako lahko pomaga zmanjšati vzorčno pristranskost, ki nastane, ker so določeni segmenti populacije v anketah premalo zastopani (npr. zaradi neodgovora ali nepokritja). Uteži
          takim segmentom iz vzorca pripišejo večjo vrednost, medtem ko preveč zastopanim pripišejo manjši pomen."),
      easyClose = TRUE,
      size = "l",
      footer = modalButton(icon("xmark"))
    ))
  })
  
  # show notification about duration of use only online (shinyapps.io)
  is_online_server <- nzchar(Sys.getenv("SHINY_PORT"))
  
  observe({
    if(is_online_server){
      showNotification(ui = HTML("<h4>Opozorilo o trajanju aplikacije</h4> </br>
                               Uporabljate spletno verzijo aplikacije. Po 15 minutah neuporabe oziroma neaktivnosti
                               se bo aplikacija ustavila in vnešeni podatki bodo ponastavljeni. Priporočamo, da si shranite svoje delo. </br></br>
                               Časovno neomejena uporaba je na voljo lokalno, navodila so dostopna <a href='https://github.com/lukastrlekar/SurveyWeightingGUI' target='_blank'>tukaj (GitHub)</a>."),
                       duration = NULL,
                       closeButton = TRUE)
    }
  }) 
  
  # Data upload tab -------------------------------------------------------------
  
  raw_data <- reactive({
    req(input$upload_raw_data)
    
    df <- try(load_file(path = input$upload_raw_data$datapath,
                        ext = input$raw_data_file_type,
                        decimal_sep = input$decimal_separator,
                        delim = input$delimiter,
                        na_strings = input$na_strings),
              silent = TRUE)
    
    if(!is.data.frame(df)){
      shinyjs::show("message_wrong_raw_data")
      return(data.frame())
      
    } else {
      shinyjs::hide("message_wrong_raw_data")
      return(df)
    }
  })
  
  output$raw_data_table <- renderDT({
    req(input$upload_raw_data)
    
    if(nrow(raw_data()) != 0){
      if(input$raw_data_file_type == "sav"){
        raw_data <- try(haven::as_factor(raw_data(), levels = "both"), silent = TRUE)
        
        if(!is.data.frame(raw_data)) raw_data <- haven::as_factor(raw_data())
        
      } else {
        raw_data <- raw_data()
      }
      
      # https://stackoverflow.com/questions/58526047/customizing-how-datatables-displays-missing-values-in-shiny
      rowCallback <- c(
        "function(row, data){",
        "  for(var i=0; i<data.length; i++){",
        "    if(data[i] === null){",
        "      $('td:eq('+i+')', row).html('NA')",
        "        .css({'color': 'rgb(151,151,151)', 'font-style': 'italic'});",
        "    }",
        "  }",
        "}"  
      )
      
      datatable(data = raw_data,
                class = "nowrap cell-border hover bordered-header",
                rownames = FALSE,
                options = list(dom = "ltip",
                               scrollX = TRUE,
                               rowCallback = JS(rowCallback)))
    }
  })

  # Select variables tab -------------------------------------------------------

  observe({
    updatePickerInput(session = session,
                      inputId = "one_dim_variables",
                      choices = names(raw_data()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "two_dim_variables",
                      choices = names(raw_data()))
  })
  
  observeEvent(input$upload_raw_data, {
    shinyjs::enable(id = "add")
    shinyjs::enable(id = "download_excel_input")
    lapply(seq_len(input$add), function(i) removeUI(selector = paste0("div:has(> #two_dim_variables",i,")")))
    lapply(seq_len(input$add), function(i) session$sendCustomMessage(type = "resetValue", message = paste0("two_dim_variables", i)))
  })
  
  observeEvent(input$add, {
    insertUI(
      selector = "#add",
      where = "beforeBegin",
      ui =
        pickerInput(
          inputId = paste0("two_dim_variables",input$add),
          label = NULL,
          choices = names(raw_data()),
          options = pickerOptions(actionsBox = TRUE,
                                  liveSearch = TRUE,
                                  maxOptions = 2),
          multiple = TRUE))
  })
  
  all_two_dim_vars <- reactive({
    c(list(input$two_dim_variables),
      lapply(seq_len(input$add), function(i) input[[paste0("two_dim_variables", i)]]))
  })
  
  output$selected_variables <- renderText({
    if(all(is.null(unlist(all_two_dim_vars()))) && is.null(input$one_dim_variables)){
      return(NULL)
      
    } else if(all(is.null(unlist(all_two_dim_vars())))){
      paste("<strong>Izbrane spremenljivke:</strong> <i>", 
            paste(input$one_dim_variables, collapse = ", "),"</i>")
      
    } else if(is.null(input$one_dim_variables)){
      paste("<strong>Izbrane spremenljivke:</strong> <i>", 
            paste(sapply(all_two_dim_vars(), paste, collapse = " x "), collapse = ", "),"</i>")
      
    } else {
      paste0("<strong>Izbrane spremenljivke:</strong> <i>", 
             paste(input$one_dim_variables, collapse = ", "),", ",
             paste(sapply(all_two_dim_vars(), paste, collapse = " x "), collapse = ", "), "</i>")
    }
  })
  
  output$preview_tables <- renderUI({
    displayed_tables <- display_tables(orig_data = raw_data(),
                                       one_dimensional_raking_variables = input$one_dim_variables,
                                       two_dimensional_raking_variables = all_two_dim_vars(),
                                       drop_zero = input$drop_zero_categories,
                                       drop_zero_missing = input$drop_missings_categories)
    n_tabs <- length(displayed_tables)
    
    displayed_tabs <- lapply(seq_len(n_tabs), function(i){
      tabPanel(title = names(displayed_tables)[[i]],
               br(),
               renderTable(displayed_tables[[i]], hover = TRUE, bordered = TRUE, spacing = "xs", align = "c",
                           caption = attr(x = raw_data()[[names(displayed_tables)[i]]], which = "label", exact = TRUE),
                           caption.placement = "top"))})
    
    do.call(tabsetPanel, displayed_tabs)
  })
  

  # Excel file for margins input --------------------------------------------
  
  output$download_excel_input <- downloadHandler(
    filename = function() {
      "Vnos_margin.xlsx"
    },
    
    content = function(file) {
      if(is.null(input$one_dim_variables) && is.null(unlist(all_two_dim_vars()))){
        showModal(modalDialog(HTML("<strong><center>Ni izbrane nobene spremenljivke</center></strong>"),
                              easyClose = TRUE,
                              footer = modalButton(icon("xmark"))))
        return()
      }
      
      # Shows message when file is loading and closes it when it is finished
      showModal(modalDialog(HTML("<h3><center>Prenašanje datoteke</center></h3>"),
                            shinycssloaders::withSpinner(uiOutput("loading"), type = 8),
                            footer = NULL))
      on.exit(removeModal())
      
      input_file_output_tables_excel(orig_data = raw_data(),
                                     one_dimensional_raking_variables = input$one_dim_variables,
                                     two_dimensional_raking_variables = all_two_dim_vars(),
                                     file_name = file, 
                                     drop_zero = input$drop_zero_categories,
                                     drop_zero_missing = input$drop_missings_categories)
    })
  
  
  # Input margins in App -------------------------------------------------------
  
  inputed_tables <- reactive({
    input_tables(orig_data = raw_data(),
                 one_dimensional_raking_variables = input$one_dim_variables,
                 two_dimensional_raking_variables = all_two_dim_vars(),
                 drop_zero = input$drop_zero_categories,
                 drop_zero_missing = input$drop_missings_categories)
  })
  
  clean_names <- reactive({
    names(inputed_tables())
  })
  
  # Reset input table on active tab
  observeEvent(input$render_input_tables, {
    lapply(seq_along(inputed_tables()), function(i) {
      output[[paste0('input_table_', clean_names()[[i]])]] <- renderRHandsontable({
        rhandsontable(inputed_tables()[[i]])
      })
    })
  })
  
  # Render input tables
  output$hot_tables <- renderUI({
    lapply(seq_along(inputed_tables()), function(i) {
      output[[paste0('input_table_', clean_names()[[i]])]] <- renderRHandsontable({
        rhandsontable(inputed_tables()[[i]])
      })
    })
    
    inputed_tabs <- lapply(seq_along(inputed_tables()), function(i){
      tabPanel(title = clean_names()[[i]],
               br(),
               column(10,
                      rHandsontableOutput(paste0("input_table_", clean_names()[[i]])),
                      br(),
                      br()),
               column(2,
                      htmlOutput(paste0("input_error_header_", clean_names()[[i]])),
                      htmlOutput(paste0("input_error_messages1_", clean_names()[[i]])),
                      htmlOutput(paste0("input_error_messages2_", clean_names()[[i]])),
                      htmlOutput(paste0("input_warning_header_", clean_names()[[i]])),
                      htmlOutput(paste0("input_warning_messages_", clean_names()[[i]])),
                      htmlOutput(paste0("input_recommend_messages_", clean_names()[[i]])),
                      htmlOutput(paste0("input_sum_message_", clean_names()[[i]]))))})
    
    do.call(tabsetPanel, inputed_tabs)
  })
  
  # Calculate population margins based on user inputted values
  observe({
    lapply(seq_along(inputed_tables()), function(i) {
      if(!is.null(input[[paste0("input_table_", clean_names()[[i]])]])){
        df <- as.data.frame(hot_to_r(input[[paste0("input_table_", clean_names()[[i]])]]))
        n <- nrow(df)
        n_col <- ncol(df)
        n_miss <- NULL # create object to include posititon of potential missing rows so they can be highlighted
        
        if(n_col == 5){
          if(any(df[,1] == "Missing")){
            val_sample <- df[df[,1] == "Missing", 3]
            val_user <- df[df[,1] == "Missing", 5]
            n_miss <- which(df[,1] == "Missing") # rows with Missing category/ies
            
            df[n_miss, 4] <- 
              ifelse(val_sample == 0 & val_user == 0, 0,
                     ifelse(val_sample != 0 & val_user == 0, val_sample,
                            ifelse(val_sample == 0 & val_user != 0, val_sample,
                                   ifelse(val_sample != 0 & val_user != 0 & val_user > val_sample, val_sample,
                                          ifelse(val_sample != 0 & val_user != 0 & val_user <= val_sample, val_user, 0)))))
            
            val_miss <- sum(df[n_miss, 4])
            
            df[-c(n_miss,n),4] <- (df[-c(n_miss,n),5]*(100 - val_miss))/(sum(df[-c(n_miss,n),5]))
            
          } else{
            df[-n,4] <- (df[-n,5]/sum(df[-n,5]))*100
          } 
          
        } else if(n_col == 6){
          if(any(df[,1] == "Missing") || any(df[,2] == "Missing")){
            val_sample <- df[df[,1] == "Missing" | df[,2] == "Missing", 4]
            val_user <- df[df[,1] == "Missing" | df[,2] == "Missing", 6]
            n_miss <- which(df[,1] == "Missing" | df[,2] == "Missing") # rows with Missing category/ies
            
            df[n_miss, 5] <- 
              ifelse(val_sample == 0 & val_user == 0, 0,
                     ifelse(val_sample != 0 & val_user == 0, val_sample,
                            ifelse(val_sample == 0 & val_user != 0, val_sample,
                                   ifelse(val_sample != 0 & val_user != 0 & val_user > val_sample, val_sample,
                                          ifelse(val_sample != 0 & val_user != 0 & val_user <= val_sample, val_user, val_sample)))))
            
            val_miss <- sum(df[n_miss, 5])
            
            df[-c(n_miss,n),5] <- (df[-c(n_miss,n),6]*(100 - val_miss))/(sum(df[-c(n_miss,n),6]))
          } else{
            df[-n,5] <- (df[-n,6]/sum(df[-n,6]))*100
          } 
        }

        col_highlight <- n_col - 1
        row_highlight_error <- which((df[,n_col-1] == 0 & df[,n_col-3] > 0) | (df[,n_col-1] > 0 & df[,n_col-3] == 0))
        
        row_highlight_critical <- which(df[,n_col-1] > 0 & df[,n_col-1] < 1)
        row_highlight_critical <- row_highlight_critical[row_highlight_critical %in% n_miss == FALSE]
        
        row_highlight_small <- which(df[,n_col-1] >= 1 & df[,n_col-1] <= 5)
        row_highlight_small <- row_highlight_small[row_highlight_small %in% n_miss == FALSE]
        
        col_highlight_sample <- n_col - 3
        row_highlight_critical_sample <- which(df[,n_col-3] >= 1 & df[,n_col-3] <= 10)
        row_highlight_critical_sample <- row_highlight_critical_sample[row_highlight_critical_sample %in% n_miss == FALSE]
        
        row_highlight_small_sample <- which(df[,n_col-3] >= 11 & df[,n_col-3] <= 30)
        row_highlight_small_sample <- row_highlight_small_sample[row_highlight_small_sample %in% n_miss == FALSE]
        
        df[n,n_col-1] <- sum(df[-n,n_col-1])
        df[n,n_col] <- sum(df[-n,n_col])

        output[[paste0("input_table_", clean_names()[[i]])]] <- renderRHandsontable({
          rhandsontable(df,
                        rowHeaders = NULL,  contextMenu = FALSE, stretchH = "all",
                        # 1 must be deducted because JS start counting at 0 but R at 1
                        customBorders = list(list(
                          range = list(from = list(row = 0, col = n_col-1),
                                       to = list(row = n-2, col = n_col-1)),
                          top = list(width = 2, color = "#337ab7"),
                          left = list(width = 2, color = "#337ab7"),
                          bottom = list(width = 2, color = "#337ab7"),
                          right = list(width = 2, color = "#337ab7"))),
                        col_highlight = col_highlight - 1,
                        row_highlight_error = row_highlight_error - 1,
                        row_highlight_critical = row_highlight_critical - 1,
                        row_highlight_small = row_highlight_small - 1,
                        miss_row = n_miss - 1,
                        col_highlight_sample = col_highlight_sample - 1,
                        row_highlight_critical_sample = row_highlight_critical_sample - 1,
                        row_highlight_small_sample = row_highlight_small_sample - 1,
                        row_last = n - 1) %>%
            hot_cols(renderer =
                "function(instance, td, row, col, prop, value, cellProperties) {
                  Handsontable.renderers.NumericRenderer.apply(this, arguments);
                  
                  if (instance.params) {
                    hcols = instance.params.col_highlight
                    hcols = hcols instanceof Array ? hcols : [hcols]
                    
                    hrowserror = instance.params.row_highlight_error
                    hrowserror = hrowserror instanceof Array ? hrowserror : [hrowserror]
                    
                    hrowscritical = instance.params.row_highlight_critical
                    hrowscritical = hrowscritical instanceof Array ? hrowscritical : [hrowscritical]
                    
                    hrowssmall = instance.params.row_highlight_small
                    hrowssmall = hrowssmall instanceof Array ? hrowssmall : [hrowssmall]
                    
                    missrows = instance.params.miss_row
                    missrows = missrows instanceof Array ? missrows : [missrows]
                    
                    if (hcols.includes(col) && hrowserror.includes(row)) {
                      td.style.background = '#FF4B4B';
                    }
                    
                    if (hcols.includes(col) && hrowscritical.includes(row)) {
                      td.style.background = '#FFC000';
                    }
                    
                    if (hcols.includes(col) && hrowssmall.includes(row)) {
                      td.style.background = '#FFECAF';
                    }
                    
                    if (missrows.includes(row)) {
                      td.style.background = '#D9D9D9';
                    }
                    
                    hcolssample = instance.params.col_highlight_sample
                    hcolssample = hcolssample instanceof Array ? hcolssample : [hcolssample]
                    
                    hrowscriticalsample = instance.params.row_highlight_critical_sample
                    hrowscriticalsample = hrowscriticalsample instanceof Array ? hrowscriticalsample : [hrowscriticalsample]
                    
                    hrowssmallsample = instance.params.row_highlight_small_sample
                    hrowssmallsample = hrowssmallsample instanceof Array ? hrowssmallsample : [hrowssmallsample]
                    
                    if (hcolssample.includes(col) && hrowscriticalsample.includes(row)) {
                      td.style.background = '#FFC000';
                    }
                    
                    if (hcolssample.includes(col) && hrowssmallsample.includes(row)) {
                      td.style.background = '#FFECAF';
                    }
                    
                    lastrow = instance.params.row_last
                    lastrow = lastrow instanceof Array ? lastrow : [lastrow]
                    
                    if (lastrow.includes(row)) {
                      td.style.fontWeight = 'bold';
                    }
                  }
                }" ) %>%  
            hot_row(row = nrow(df), readOnly = TRUE) %>%
            hot_col(col = 1:(n_col-1), readOnly = TRUE) %>%
            hot_col(col = n_col-3, format = "0,0") %>% 
            hot_col(col = (n_col-1):n_col, format = "0.00000") %>% 
            hot_validate_numeric(cols = n_col, min = 0, max = 100)
        })
        
        output[[paste0("input_error_header_", clean_names()[[i]])]] <- renderText({
          if(length(row_highlight_error) != 0){
            paste('<strong>Kritična opozorila (uteževanje ne bo delovalo):</strong></br>')
          }
        })
        
        ### Warning: Error in if: missing value where TRUE/FALSE needed
        output[[paste0("input_error_messages1_", clean_names()[[i]])]] <- renderText({
          if(isTRUE(length(row_highlight_error) != 0 && any((df[-c(n_miss,n),n_col-1] == 0 & df[-c(n_miss,n),n_col-3] > 0)))){
            paste('<p style="background-color:#FF4B4B;">Ničelna populacijska in neničelna vzorčna margina</p>')
          }
        })
        
        output[[paste0("input_error_messages2_", clean_names()[[i]])]] <- renderText({
          if(length(row_highlight_error) != 0 && any((df[-c(n_miss,n),n_col-1] > 0 & df[-c(n_miss,n),n_col-3] == 0))){
            paste('<p style="background-color:#FF4B4B;">Neničelna populacijska in ničelna vzorčna margina</p>')
          }
        })
        
        output[[paste0("input_warning_header_", clean_names()[[i]])]] <- renderText({
          if(any(length(row_highlight_critical) != 0,
                 length(row_highlight_critical_sample) != 0,
                 length(row_highlight_small) != 0,
                 length(row_highlight_small_sample) != 0)){
            paste('<br/><strong>Priporočila: </br>
                  <small>(priporočena je priključitev obarvanih kategorij k vsebinsko podobni kategoriji)</small></strong></br>')
          }
        })
        
        output[[paste0("input_warning_messages_", clean_names()[[i]])]] <- renderText({
          if(length(row_highlight_critical) != 0 && length(row_highlight_critical_sample) != 0){
            paste('<p style="background-color:#FFC000;">Kritično majhne kategorije (n ≤ 10) in populacijske margine (< 1%)</p>')
          } else if(length(row_highlight_critical) != 0 && length(row_highlight_critical_sample) == 0){
            paste('<p style="background-color:#FFC000;">Kritično majhne populacijske margine (< 1%)</p>')
          } else if(length(row_highlight_critical) == 0 && length(row_highlight_critical_sample) != 0){
            paste('<p style="background-color:#FFC000;">Kritično majhne kategorije (n ≤ 10)</p>')
          }
        })
        
        output[[paste0("input_recommend_messages_", clean_names()[[i]])]] <- renderText({
          if(length(row_highlight_small) != 0 && length(row_highlight_small_sample) != 0){
            paste('<p style="background-color:#FFECAF;">Majhne kategorije (n ≤ 30) in populacijske margine (< 5%)</p>')
          } else if(length(row_highlight_small) != 0 && length(row_highlight_small_sample) == 0){
            paste('<p style="background-color:#FFECAF;">Majhne populacijske margine (< 5%)</p>')
          } else if(length(row_highlight_small) == 0 && length(row_highlight_small_sample) != 0){
            paste('<p style="background-color:#FFECAF;">Majhne kategorije (n ≤ 30)</p>')
          }
        })
        
        output[[paste0("input_sum_message_", clean_names()[[i]])]] <- renderText({
          s <- df[n,n_col]
          if((s != 1 && s != 100) && s != 0){
            paste('<br/><strong>Opomba:</strong></br>
                  <p style="background-color:#FFFFC1;">Vsota vnesenih margin ni 100%</br><small>(margine bodo preračunane, da bo vsota 100%)</small></p>')
          }
        })
      }
    })
  })
  
  output$download_inputed_tables <- downloadHandler(
    filename = function() {
      "Vnesene_margine.RData"
    },
    
    content = function(file) {
      if(is.null(input$one_dim_variables) && is.null(unlist(all_two_dim_vars()))){
        showModal(modalDialog(HTML("<strong><center>Ni izbrane nobene spremenljivke</center></strong>"),
                              easyClose = TRUE,
                              footer = modalButton(icon("xmark"))))
        return()
      }
      
      inputted_margins_list <- lapply(seq_along(inputed_tables()), function(i) {
        hot_to_r(input[[paste0("input_table_", clean_names()[[i]])]])
      })
      
      names(inputted_margins_list) <- clean_names()
      
      save(object = inputted_margins_list, file = file)
    })

  # Weighting tab ---------------------------------------------------------------
    
  output$which_weighting_vars <- renderUI({
    input_true <- sapply(seq_len(length(inputed_tables())), function(i) {
      is.null(input[[paste0("input_table_", clean_names()[[i]])]])
    }) # TOLE PREGLEJ
    
    file_input <- fileInput(inputId = "upload_margins_data", 
                            label = "Naloži datoteko z vnešenimi populacijskimi marginami",
                            multiple = FALSE,
                            accept = c(".xlsx", ".RData"),
                            buttonLabel = "Naloži datoteko")
    
    if(all(input_true)) {
      file_input
      
    } else {
      tagList(
        radioButtons("radio_which_weighting_vars", 
                     label = "Katere populacijske (ciljne) margine želite upoštevati?",
                     choices = c("Na novo vnešene v zavihku Vnos margin v aplikaciji" = "newly_inputed",
                                 "Shranjene od prej (v Excel ali RData datoteki)" = "saved_inputed")),
        conditionalPanel(condition = "input.radio_which_weighting_vars == 'saved_inputed'",
                         file_input)
      )
    } 
  })
  
  margins_data <- reactive({
    if(isTRUE(input$radio_which_weighting_vars == "newly_inputed")){
      sheet_list <- lapply(seq_len(length(inputed_tables())), function(i) {
        hot_to_r(input[[paste0("input_table_", clean_names()[[i]])]])
      })
      
      names(sheet_list) <- clean_names()
      return(list(sheet_names = names(sheet_list),
                  sheet_list = sheet_list))
      
    } else if(!is.null(input$upload_margins_data) || isTRUE(input$radio_which_weighting_vars == "saved_inputed")){
      req(input$upload_margins_data)
      
      ext <- tools::file_ext(input$upload_margins_data$name)
      
      if(!(ext %in% c("xlsx", "RData"))){
        shinyjs::show("message_wrong_file")
      } else if(is.null(input$upload_raw_data)){
        shinyjs::show("message_no_file")
      } else {
        shinyjs::hide("message_wrong_file")
        shinyjs::hide("message_no_file")
      }
      
      sheet_names <- NULL
      sheet_list <- NULL
      
      if(ext == "xlsx"){
        sheet_names <- openxlsx::getSheetNames(input$upload_margins_data$datapath)
        sheet_list <- lapply(sheet_names, function(sn){openxlsx::read.xlsx(input$upload_margins_data$datapath, sheet = sn)})
        
        file_names_indicator <- vapply(sheet_list, FUN = function(x){
          if(!all(c("Frekvenca", "Populacijski.(ciljni).%") %in% names(x)[1:5])){
            TRUE
          } else {
            FALSE
          }
        }, FUN.VALUE = logical(1))
        
        if(any(file_names_indicator)){
          shinyjs::show("message_wrong_file_structure")
          sheet_names <- NULL
          sheet_list <- NULL
        } else {
          shinyjs::hide("message_wrong_file_structure")
          
          # ensure that full variable names are shown, because Excel sheet names are limited to 31 characters
          variable_names <- vapply(sheet_list, FUN = function(x){
            var_names <- names(x)[1:2][names(x)[1:2] != "Frekvenca"]
            if(length(var_names) == 2){
              paste0(var_names, collapse = " x ")
            } else {
              var_names
            }
          }, FUN.VALUE = character(1))
          
          # Check if variables are present in uploaded data
          warning_counter <- vapply(sheet_list, FUN = function(x){
            var_names <- names(x)[1:2][names(x)[1:2] != "Frekvenca"]
            if(all(var_names %in% names(raw_data()))){
              FALSE
            } else {
              TRUE
            }
          }, FUN.VALUE = logical(1))
          
          # Show error if no variables are present in uploaded data
          if(all(warning_counter)){
            shinyjs::show("message_no_variables_present")
          } else if(any(warning_counter)) {
          # Show warning if some variables are not present in uploaded data
            shinyjs::show("message_some_variables_not_present")
          } else {
            shinyjs::hide("message_no_variables_present")
            shinyjs::hide("message_some_variables_not_present")
          }
          
          variable_names <- variable_names[!warning_counter]
          sheet_list <- sheet_list[!warning_counter]
          names(sheet_list) <- variable_names
          sheet_names <- variable_names
        }
        
      } else if(ext == "RData"){
        # TODO preveri strukturo in ... kot pri xlsx
        sheet_list <- load_to_environment(input$upload_margins_data$datapath)
        sheet_names <- names(sheet_list)
      }
      
      return(list(sheet_names = sheet_names,
                  sheet_list = sheet_list))
    } 
  })
  
  # Weighting variables selection
  observe({
    updatePickerInput(session = session,
                      inputId = "weights_variables",
                      choices = margins_data()$sheet_names,
                      selected = margins_data()$sheet_names)
  })
  
  # Case id variable selection
  observe({
    updatePickerInput(session = session,
                      inputId = "case_id_variable", choices = names(raw_data()))
  })
  
  # Base weights variable selection
  observe({
    updatePickerInput(session = session,
                      inputId = "base_weights_variable",
                      choices = names(raw_data()))
  })
  
  observe({
    req(input$base_weights_variable)
    feedbackDanger("base_weights_variable", anyNA(raw_data()[[input$base_weights_variable]]) || !is.numeric(raw_data()[[input$base_weights_variable]]),
                   "Uteži ne smejo vsebovati manjkajočih ali neštevilskih vrednosti", color = "red", icon = NULL)
  })
  
  observe({
    feedbackDanger("cap_weights", is.na(input$cap_weights) || input$cap_weights <= 1, "Neveljavna vrednost", color = "red", icon = NULL)
  })

  observe({
    feedbackDanger("min_weight", is.na(input$min_weight) || input$min_weight < 0 || input$min_weight >= 1, "Neveljavna vrednost", color = "red", icon = NULL)
  })
  
  observe({
    feedbackDanger("max_weight", is.na(input$max_weight) || input$max_weight <= 1, "Neveljavna vrednost", color = "red", icon = NULL)
  })

  observe({
    if(input$package == "anesrake"){
      shinyjs::enable(id = "cut_weights_iterative")
    } else {
      updateCheckboxInput(inputId = "cut_weights_iterative", value = FALSE)
      shinyjs::disable(id = "cut_weights_iterative")
    }
  })
  
  # Perform weighting when button `run_weighting` is clicked
  weighting_output <- eventReactive(input$run_weighting, {
    if(input$cut_weights_iterative == TRUE){
      if(is.na(input$cap_weights) || input$cap_weights <= 1){
        cap <- 999999
      } else {
        cap <- input$cap_weights
      }
    } else {cap <- 999999}
    
    if(input$cut_weights_after == TRUE){
      if(is.na(input$min_weight) || input$min_weight < 0 || input$min_weight >= 1){
        lower <- -Inf
      } else {
        lower <- input$min_weight
      }
      
      if(is.na(input$max_weight) || input$max_weight <= 1){
        upper <- Inf
      } else {
        upper <- input$max_weight
      }
      
    } else {
      lower <- -Inf
      upper <- Inf
    }
    
    if(input$convergence_input == TRUE){
      convcrit <- input$convergence_criterion
      epsilon <- input$convergence_criterion_epsilon
    } else {
      convcrit <- 0.01
      epsilon <- 1e-10
      }
    
    if(input$base_weights_selection == TRUE){
      if(is.null(input$base_weights_variable)){
        shinyjs::show("message_no_base_weights_selected")
        weightvec <- NULL
      } else {
        shinyjs::hide("message_no_base_weights_selected")
        weightvec <- raw_data()[[input$base_weights_variable]]
        
        if(anyNA(weightvec) || !is.numeric(weightvec)){
          weightvec <- NULL
        }
      }
      
    } else {
      shinyjs::hide("message_no_base_weights_selected")
      weightvec <- NULL
    }
    
    if(input$case_id_selection == TRUE){
      case_id <- input$case_id_variable
    } else {
      case_id <- NULL
    }
    
    perform_weighting(orig_data = raw_data(),
                      margins_data = margins_data(),
                      all_raking_variables = input$weights_variables,
                      case_id = case_id,
                      lower = lower,
                      upper = upper,
                      cap = cap, 
                      convcrit = convcrit,
                      epsilon = epsilon,
                      weightvec = weightvec,
                      package = input$package)
  })

  cut_weights_text <- eventReactive(input$run_weighting, {
    text1 <- ifelse(input$cut_weights_after == FALSE || is.na(input$min_weight) || input$min_weight < 0 || input$min_weight >= 1, NA, input$min_weight)
    text2 <- ifelse(input$cut_weights_after == FALSE || is.na(input$max_weight) || input$max_weight <= 1, NA, input$max_weight)
    text <- c(text1, text2)
    
    show_text <- if(all(is.na(text))){
      "IZKLOPLJENO"
    } else if(all(!is.na(text))) {
      paste0("VKLOPLJENO (minimalna vrednost uteži: ", input$min_weight, ", maksimalna vrednost uteži: ", input$max_weight, ")")
    } else if(is.na(text)[1]) {
      paste0("IZKLOPLJENO za minimalno vrednost, VKLOPLJENO za maksimalno vrednost (maksimalna vrednost uteži: ", input$max_weight, ")")
    } else if(is.na(text)[2]) {
      paste0("IZKLOPLJENO za maksimalno vrednost, VKLOPLJENO za minimalno vrednost (minimalna vrednost uteži: ", input$min_weight, ")")
    }
    
    return(list(text_cut = show_text,
                text_cut_iter = ifelse(input$cut_weights_iterative == FALSE || is.na(input$cap_weights) || input$cap_weights < 1, "IZKLOPLJENO", paste0("VKLOPLJENO (maksimalna vrednost uteži: ", input$cap_weights,")"))))
  })
  
  # Show weighting summary
  output$code_summary <- renderText({
    req(input$run_weighting)
    
    summ <- try(summary(weighting_output()), silent = TRUE)
    deff <- design_effect(weighting_output()$weightvec)
    
    HTML(paste("<b>Paket za raking:</b>", weighting_output()$package, "</br></br>"),
         paste("<b>Konvergenca:</b>", summ$convergence, "</br></br>"),
         paste("<b>Iterativno rezanje uteži:</b>", cut_weights_text()$text_cut_iter, "</br></br>"),
         paste("<b>Rezanje uteži po koncu iteracij:</b>", cut_weights_text()$text_cut, "</br></br>"),
         paste("<b>Uteževalne spremenljivke:</b>", paste0(summ$raking.variables, collapse = ", "), "</br></br>"),
         paste("<b>Privzete uteži:</b>", summ$base.weights, "</br></br>"),
         paste0("<b>Vzorčni učinek (design effect): </b>", round(deff, 5), ". Porast vzorčne variance zaradi uteževanja: ", round((deff-1)*100,2), "% ",
                '(<a href="https://www.proquest.com/openview/8196b6985339ec9885193d56a5c44ca0/1?pq-origsite=gscholar&cbl=105444" target="_blank">Kish (1992)</a>;
                <a href="https://www.hzu.edu.in/uploads/2020/9/Applied%20Survey%20Data%20Analysis%20(Chapman%20&%20Hall%20CRC%20Statistics%20in%20the%20Social%20and%20Behavioral%20Scie).pdf" target="_blank">Heeringa in drugi (2010)</a>, str. 44-45).'))
  })

  output$weights_table <- renderDT({
    datatable(data = setNames(data.frame(caseid = weighting_output()$caseid,
                                         weights = weighting_output()$weightvec),
                              nm = c(weighting_output()$caseid_name, "weights")),
              rownames= FALSE,
              class = "compact row-border hover",
              options = list(dom = "t",
                             paging = FALSE,
                             scrollY = "500px",
                             scrollCollapse = TRUE))
  })
  
  # Show weighting variables frequency tables
  observe({
    output$weighting_variables_tables <- renderUI({
      sheet_names <- margins_data()$sheet_names
      sheet_names <- sheet_names[sheet_names %in% input$weights_variables]
      
      sheet_list <- margins_data()$sheet_list
      sheet_list <- sheet_list[sheet_names]
      
      n_tabs <- length(sheet_names)
      
      if(input$base_weights_selection == TRUE){
        if(is.null(input$base_weights_variable)){
          prevec <- NULL
        } else {
          prevec <- raw_data()[[input$base_weights_variable]]
          
          if(anyNA(prevec) || !is.numeric(prevec)){
            prevec <- NULL
          }
        }
      } else {
        prevec <- NULL
      }
      
      displayed_tabs <- lapply(seq_len(n_tabs), function(i){
        tabPanel(title = sheet_names[i],
                 br(),
                 datatable(display_tables_weighting_vars(orig_data = raw_data(),
                                                         sheet_list_table = sheet_list[[i]],
                                                         prevec = prevec),
                           rownames = FALSE,
                           class = "compact cell-border hover bordered-header",
                           options = list(dom = "t",
                                          paging = FALSE,
                                          ordering = FALSE,
                                          columnDefs = list(list(className = 'dt-center', targets = "_all")))) %>% 
                   formatRound(columns = 3:4, digits = 2) %>% 
                   formatRound(columns = 2, digits = 0),
                 br()
        )})
      do.call(tabsetPanel, displayed_tabs)
    })
  })
  
  observeEvent(input$run_weighting, {
    output$weighting_variables_tables <- renderUI({
      tbl_list <- weightassess(inputter = weighting_output()$targets,
                               dataframe = weighting_output()$dataframe,
                               weightvec = weighting_output()$weightvec,
                               prevec = weighting_output()$prevec)
      n_tabs <- length(tbl_list)
      
      displayed_tabs <- lapply(seq_len(n_tabs), function(i){
        tabPanel(title = weighting_output()$varsused[i],
                 br(),
                 datatable(tbl_list[[i]],
                           rownames = FALSE,
                           class = "compact cell-border hover bordered-header",
                           options = list(dom = "t",
                                          paging = FALSE,
                                          ordering = FALSE,
                                          columnDefs = list(list(className = 'dt-center', targets = "_all")))) %>% 
                   formatRound(columns = c(3,5:7), digits = 2) %>% 
                   formatRound(columns = 8, digits = 5) %>% 
                   formatRound(columns = c(2,4), digits = 0),
                 br()
        )})
      do.call(tabsetPanel, displayed_tabs)
    })
  })
  
  # Show distribution of weights
  output$weights_plot <- renderPlot({
    req(input$run_weighting)
    vec <- weighting_output()$weightvec
    
    plot(hist(x = vec,
              xlim = c(0, max(vec) + 1),
              breaks = seq(min(vec), max(vec), by = ((max(vec) - min(vec))/(length(vec))))),
         xlab = "Uteži", ylab = "Frekvence",
         main = "Porazdelitev uteži")
  })
  
  # Show weights' descriptive statistics
  output$summary_stat <- renderTable({
    req(input$run_weighting)
    data.frame(t(unclass(summary(weighting_output()$weightvec))), check.names = FALSE)
  },
  caption =  "<b><span style='color: #333'> Opisne statistike uteži </b>",
  caption.placement = "top")
  
  # Download weights file
  output$jump_to_download_weights_ui <- renderUI({
    if(!is.null(weighting_output())){
      tagList(
        actionButton("jump_to_download_weights",
                     label = HTML("&nbspPojdi na prenos uteži"),
                     icon = shiny::icon("arrow-up-right-from-square"),
                     width = "100%",
                     class = "btn-primary"),
        br(),
        br())
    }
  })
  
  observeEvent(input$jump_to_download_weights, {
    updateTabsetPanel(session = session, inputId = "weighting_panel", selected = "download_weights_tab")
  })
  
  output$download_weights <- downloadHandler(
    filename = function() {
      paste0("Utezi_", tools::file_path_sans_ext(input$upload_raw_data$name), "_", paste0(weighting_output()$varsused, collapse = "_"), ".", input$select_weights_file_type)
    },
    content = function(file) {
      download_weights(weights_object = weighting_output(),
                       file_type = input$select_weights_file_type,
                       file_name = file,
                       separator = input$delimiter_weights,
                       decimal = input$decimal_separator_weights,
                       quote_col_names = input$quote_col_names_weights)
    })
  
  # Download diagnostic file 
  output$weighting_diagnostic_download_ui <- renderUI({
    if(!is.null(weighting_output())){
      downloadButton("diagnostic_file_download",
                     label = HTML("&nbspPrenesi povzetek uteževanja"),
                     icon = shiny::icon("file-excel"),
                     class = "btn-secondary",
                     style = "width: 100%;")
    }
  })
  
  output$diagnostic_file_download <- downloadHandler(
    filename = function() {
      "Povzetek_utezevanja.xlsx"
    },
    content = function(file) {
      download_weighting_diagnostic(weights_object = weighting_output(),
                                    file_name = file,
                                    cut_text = cut_weights_text()$text_cut,
                                    iter_cut_text = cut_weights_text()$text_cut_iter,
                                    is_online_server = is_online_server)
    })
  
  # Download survey design object
  output$download_survey_design <- downloadHandler(
    filename = function() {
      paste0("svydesign_", tools::file_path_sans_ext(input$upload_raw_data$name), "_", paste0(weighting_output()$varsused, collapse = "_"), ".RData")
    },
    content = function(file) {
      survey_design_raking <- weighting_output()$survey_design
      save(object = survey_design_raking, file = file)
    })
  
  # Analyses tab ----------------------------------------------------------------
  
  observe({
    updatePickerInput(session = session,
                      inputId = "numeric_variables", choices = names(raw_data()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "factor_variables", choices = names(raw_data()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "analyses_weight_variable", choices = names(raw_data()))
  })
  
  # enable survey SE calculation when svydesign is present
  observe({
    if(weighting_output()$package == "survey"){
      shinyjs::enable(id = "se_calculation")
    } else {
      updateRadioButtons(inputId = "se_calculation", selected = "taylor_se")
      shinyjs::disable(id = "se_calculation")
    }
  })
  
  # vector of weights (uploaded or from raking in previous step)
  weights_vector <- reactive({
    if(!is.null(input$analyses_weight_variable) && input$select_which_weights == "included_weights"){
      weights <- labelled::user_na_to_na(raw_data()[[input$analyses_weight_variable]])
      
      validate(need((sum(!is.na(weights)) == nrow(raw_data())) && is.numeric(weights), message = FALSE))
      
      return(weights)
      
    } else {
      return(weighting_output()$weightvec)
    }
  })
  
  # Display warning/error messages for selected weights variable
  output$weights_message <- renderText({
    if(input$run_numeric_variables != 0 || input$run_factor_variables != 0){
      if(!is.null(input$analyses_weight_variable) && input$select_which_weights == "included_weights"){
        weights <- labelled::user_na_to_na(raw_data()[[input$analyses_weight_variable]])
        
        if(anyNA(weights)){
          validate("Izbrana spremenljivka uteži ne sme vsebovati manjkajočih vrednosti.")
        }
        
        if(!is.numeric(weights)){
          validate("Izbrana spremenljivka za uteži ni številska. Izračun uteženih statistik ni mogoč.")
        }
        
        if(!isTRUE(all.equal(mean(weights, na.rm = TRUE), 1))){
          # TODO dodaj postratifikacisjke uteži imajo načeloma povprečje 1
          validate("Povprečje uteži ni enako 1 (uteži niso normirane). Za analize bodo uteži normirane, da bo povprečje enako 1.")
        }
        
      } else if(is.null(input$analyses_weight_variable) && input$select_which_weights == "included_weights"){
        validate("Izbrana ni bila nobena utež.")
        
      } else if(input$select_which_weights == "new_weights" && input$run_weighting==0){
        validate("Ni prisotnih uteži.")
      }
    }
  })
  
  ### Numeric variables
  weighted_numeric_table <- eventReactive(input$run_numeric_variables, {
    req(input$numeric_variables)
    
    if(input$adjust_p_values == 0){
      p_adjust_method <- NULL
    } else {
      p_adjust_method <- input$select_p_adjust_method
    }
    
    weighted_numeric_statistics(numeric_variables = input$numeric_variables,
                                orig_data = raw_data(),
                                weights = weights_vector(),
                                p_adjust_method = p_adjust_method,
                                se_calculation = input$se_calculation,
                                survey_design = weighting_output()$survey_design)
  })
  
  # Display message about removal of non-numeric variables when calculating descriptive statistics
  output$message_numeric <- renderText({
    req(input$run_numeric_variables)
    
    if((length(weighted_numeric_table()$non_numeric_vars) == 0) && (length(weighted_numeric_table()$non_calculated_vars) == 0)){
      return(NULL)
    } else if((length(weighted_numeric_table()$non_numeric_vars) != 0) && (length(weighted_numeric_table()$non_calculated_vars) == 0)){
      paste("<small>Spremenljivke", paste(weighted_numeric_table()$non_numeric_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker vsebujejo neštevilske vrednosti.</small> <hr/>")
    } else if((length(weighted_numeric_table()$non_calculated_vars) != 0) && (length(weighted_numeric_table()$non_numeric_vars) == 0)){
      paste("<small>Spremenljivke", paste(weighted_numeric_table()$non_calculated_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker imajo ničelno varianco in izračun testne statistike ni bil mogoč.</small> <hr/>")
    } else {
      paste("<small>Spremenljivke", paste(weighted_numeric_table()$non_numeric_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker vsebujejo neštevilske vrednosti.<br/><br/>
            Spremenljivke", paste(weighted_numeric_table()$non_calculated_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker imajo ničelno varianco in izračun testne statistike ni bil mogoč.</small> <hr/>")
    }
  })
  
  output$analyses_numeric_tables <- renderUI({
    if(!is.null(weighted_numeric_table()$calculated_table)){
      n_col <- ncol(weighted_numeric_table()$calculated_table)
      
      tagList(
        p(HTML("<small>Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001</small>")),
        datatable(weighted_numeric_table()$calculated_table,
                  rownames = FALSE,
                  class = "compact cell-border hover bordered-header",
                  options = list(dom = "t",
                                 paging = FALSE,
                                 columnDefs = list(list(targets = 0:(n_col-1),
                                                        className = "dt-center")))) %>%
          formatRound(columns = (n_col-6):(n_col-1), digits = 2) %>% 
          formatRound(columns = c((n_col-8):(n_col-7), (n_col-3)), digits = 0),
        br())
    }
  })
  
  ### Categorical (factor) variables
  weighted_factor_tables <- eventReactive(input$run_factor_variables, {
    req(input$factor_variables)
    
    n_tabs <- length(input$factor_variables)
    
    if(input$adjust_p_values == 0){
      p_adjust_method <- NULL
    } else {
      p_adjust_method <- input$select_p_adjust_method
    }
    
    tables <- lapply(seq_len(n_tabs), function(i){
      create_w_table(orig_data = raw_data(),
                     variable = input$factor_variables[[i]],
                     weights = weights_vector(),
                     p_adjust_method = p_adjust_method,
                     se_calculation = input$se_calculation,
                     survey_design = weighting_output()$survey_design)
    })
    
    names(tables) <- input$factor_variables
    return(tables)
  })
  
  # Display message about removal of constant factor variables
  output$message_factor <- renderText({
    req(input$run_factor_variables)
    
    invalid_vars <- unlist(lapply(weighted_factor_tables(), "[[", "invalid_var"))
    constant_vars <- unlist(lapply(weighted_factor_tables(), "[[", "constant_var"))
    
    if((length(invalid_vars) == 0) && (length(constant_vars) == 0)){
      return(NULL)
    } else if((length(invalid_vars) != 0) && (length(constant_vars) == 0)){
      paste("<small>Spremenljivke", paste(invalid_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker nimajo veljavnih (nemanjkajočih) vrednosti.</small> <hr/>")
    } else if((length(constant_vars) != 0) && (length(invalid_vars) == 0)){
      paste("<small>Spremenljivke", paste(constant_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker so konstante (imajo le 1 veljavno kategorijo).</small> <hr/>")
    } else {
      paste("<small>Spremenljivke", paste(invalid_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker nimajo veljavnih (nemanjkajočih) vrednosti.<br/><br/>
            Spremenljivke", paste(constant_vars, collapse = ", "),
            "so bile pri izračunih izpuščene, ker so konstante (imajo le 1 veljavno kategorijo).</small> <hr/>")
    }
  })
  
  output$analyses_factor_tables <- renderUI({
    index <- vapply(weighted_factor_tables(), function(x) is.data.frame(x[["calculated_table"]]), FUN.VALUE = logical(1))
    weighted_factor_tables <- weighted_factor_tables()[index]
    
    warning_indicator <- any(unlist(lapply(weighted_factor_tables, "[[", "warning_indicator")), na.rm = TRUE)
    
    tagList(
      p(HTML("<small>Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001</small>")),
      if(warning_indicator) p(HTML("<small>! Število enot v celici je premajhno za zanesljivo oceno p vrednosti (np ≤ 5 ali n(1-p) ≤ 5).</small>")),
      
      lapply(seq_len(length(weighted_factor_tables)), function(i){
        tbl <- weighted_factor_tables[[i]][["calculated_table"]]
        n_col <- ncol(tbl)
        
        tagList(
          datatable(tbl,
                    rownames= FALSE,
                    class = "compact cell-border hover bordered-header",
                    caption = attr(x = raw_data()[[names(weighted_factor_tables)[i]]], which = "label", exact = TRUE),
                    options = list(dom = "t",
                                   paging = FALSE,
                                   columnDefs = list(list(targets = 0:(n_col-1),
                                                          className = "dt-center")))) %>%
            formatRound(columns = (n_col-8):(n_col-3), digits = 0) %>% 
            formatRound(columns = (n_col-2):(n_col-1), digits = 2),
          br())
      })
    )
  })
  
  
  ### Download analyses numeric tables
  output$download_numeric_analyses_ui <- renderUI({
    if(!is.null(weighted_numeric_table()$calculated_table)){
      downloadButton("numeric_analyses_file_download",
                     label = HTML("&nbspPrenesi tabelo opisnih statistik"),
                     icon = shiny::icon("file-excel"),
                     class = "btn-secondary")
    }
  })
  
  output$numeric_analyses_file_download <- downloadHandler(
    filename = function() {
      "Analiza_utezevanje_opisne_statistike.xlsx"
    },
    content = function(file) {
      download_analyses_numeric_table(numeric_table = weighted_numeric_table()$calculated_table,
                                      file = file)
    })
  
  ### Download analyses factor tables
  output$download_factor_analyses_ui <- renderUI({
    if(!is.null(weighted_factor_tables())){
      downloadButton("factor_analyses_file_download",
                     label = HTML("&nbspPrenesi frekvenčne tabele"),
                     icon = shiny::icon("file-excel"),
                     class = "btn-secondary")
    }
  })
  
  output$factor_analyses_file_download <- downloadHandler(
    filename = function() {
      "Analiza_utezevanje_frekvencne_tabele.xlsx"
    },
    content = function(file) {
      showModal(modalDialog(HTML("<h3><center>Prenašanje datoteke</center></h3>"),
                            shinycssloaders::withSpinner(uiOutput("loading"), type = 8),
                            footer = NULL))
      on.exit(removeModal())
      
      index <- vapply(weighted_factor_tables(), function(x) is.data.frame(x[["calculated_table"]]), FUN.VALUE = logical(1))
      weighted_factor_tables <- weighted_factor_tables()[index]
      
      download_analyses_factor_tables(
        factor_tables = lapply(weighted_factor_tables, "[[", "calculated_table"),
        orig_data = raw_data(),
        file = file,
        warning_indicator = any(unlist(lapply(weighted_factor_tables, "[[", "warning_indicator")), na.rm = TRUE))
    })
}