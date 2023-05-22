# create one-dimensional contingency table to excel to input population margins
one_dim_excel <- function(orig_data, one_dim_raking_var, wb, drop_zero, drop_zero_missing){
  temp_var <- one_dim_raking_var
  one_dim_raking_var <- substr(one_dim_raking_var, 1, 31) # Max length is 31 characters for Excel sheet name
  
  temp_df <- data.frame(table(clean_data(orig_data[[temp_var]]), useNA = "no"))
  temp_df$prop <- (temp_df$Freq/sum(temp_df$Freq))*100
  temp_df$popul <- rep(0, nrow(temp_df))
  temp_df$input <- rep(0, nrow(temp_df))
  temp_df <- rbind(temp_df, data.frame(Var1 = "Skupaj", t(colSums(temp_df[,-1, drop = FALSE]))))
  colnames(temp_df) <- c(temp_var, "Frekvenca", "Vzorcni %", "Populacijski (ciljni) %", "Vnos populacijskih margin (%)")
  
  if(drop_zero == TRUE){ # drop empty content categories
    temp_df <- temp_df[temp_df$Frekvenca > 0 | temp_df[[temp_var]] == "Missing", ,drop = FALSE]
  }
  
  if(drop_zero_missing == TRUE){ # drop emtpy missing categories
    temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df$Frekvenca == 0 & temp_df[[temp_var]] != "Missing"), ,drop = FALSE]
  }
  
  n_row <- nrow(temp_df) + 1
  n_col <- ncol(temp_df)
  
  addWorksheet(wb = wb, sheetName = one_dim_raking_var, gridLines = FALSE)
  writeData(wb = wb,
            sheet = one_dim_raking_var,
            x = temp_df,
            colNames = TRUE, rowNames = FALSE,
            borders = "all",
            headerStyle = createStyle(halign = "center", textDecoration = "bold", border = "TopBottomLeftRight"))
  addStyle(wb = wb,
           sheet = one_dim_raking_var,
           rows = 1, cols = n_col,
           style = createStyle(fontColour = "#FFFFFF", fgFill = "#000000"),
           stack = TRUE)
  
  not_miss_rows <- NULL
  
  for(j in 2:n_row){
    if(temp_df[[1]][j-1] == "Missing"){ # Missing category
      writeFormula(wb = wb, 
                   sheet = one_dim_raking_var, 
                   x = paste0("IF(AND(C",n_row-1,"=0,E",n_row-1,"=0),0,IF(AND(C",n_row-1,"<>0,E",n_row-1,"=0),C",n_row-1,",IF(AND(C",n_row-1,"=0,E",n_row-1,"<>0),C",n_row-1,",IF(AND(C",n_row-1,"<>0,E",n_row-1,"<>0,E",n_row-1,">C",n_row-1,"),C",n_row-1,",IF(AND(C",n_row-1,"<>0,E",n_row-1,"<>0,E",n_row-1,"<=C",n_row-1,"),E",n_row-1,")))))"), 
                   startCol = n_col-1, startRow = j)
      
      addStyle(wb = wb, sheet = one_dim_raking_var, 
               style = createStyle(fgFill = "#D9D9D9"), 
               rows = j, cols = 1:n_col, stack = TRUE)
      
      addStyle(wb = wb, sheet = one_dim_raking_var, 
               style = createStyle(locked = FALSE,
                                   textDecoration = "bold"), 
               rows = j, cols = n_col, stack = TRUE)
      
    } else if(temp_df[[1]][j-1] == "Skupaj"){ # Total category
      writeFormula(wb = wb, 
                   sheet = one_dim_raking_var, 
                   x = paste0("SUM(D2:D",n_row-1,")"), 
                   startCol = n_col-1, startRow = j)
      writeFormula(wb = wb, 
                   sheet = one_dim_raking_var, 
                   x = paste0("SUM(E2:E",n_row-1,")"), 
                   startCol = n_col, startRow = j)
      
      addStyle(wb = wb, sheet = one_dim_raking_var, 
               style = createStyle(border = "top", 
                                   borderStyle = "double", 
                                   textDecoration = "bold"), 
               rows = j, cols = 1:n_col, stack = TRUE)
    } else{ # content categories
      if(any(temp_df[[1]] == "Missing")){
        writeFormula(wb = wb, 
                     sheet = one_dim_raking_var, 
                     x = paste0("IFERROR(E",j,"*(100-$D$",n_row-1,")/(SUM($E$2:$E$",n_row-2,")),0)"), 
                     startCol = n_col-1, startRow = j)
      } else{
        writeFormula(wb = wb, 
                     sheet = one_dim_raking_var, 
                     x = paste0("IFERROR((E",j,"/(SUM($E$2:$E$",n_row-1,"))*100),0)"), 
                     startCol = n_col-1, startRow = j)
      }
      
      writeFormula(wb = wb, sheet = one_dim_raking_var,
                   x = paste0('IF(AND(B',j,'>=1,B',j,'<=4),"vzorec1-4",IF(AND(B',j,'>=5,B',j,'<=9),"vzorec5-9","ok"))&"_"&IF(AND(D',j,'=0,B',j,'>0,$E$',n_row,'<>0),"pop0",IF(AND(B',j,'=0,D',j,'>0,$E$',n_row,'<>0),"vzorec0",IF(AND(D',j,'>0,D',j,'<1),"vzorec1-4",IF(AND(D',j,'>=1,D',j,'<5),"vzorec5-9","ok"))))'),
                   startCol = n_col+1, startRow = j)
      
      not_miss_rows <- c(not_miss_rows, j)
    }
  }
  
  addStyle(wb = wb, 
           sheet = one_dim_raking_var, 
           style = createStyle(fontColour = "#FFFFFF"), 
           rows = not_miss_rows, cols = n_col+1, stack = TRUE)
  
  addStyle(wb = wb, 
           sheet = one_dim_raking_var, 
           style = createStyle(locked = FALSE,
                               textDecoration = "bold"), 
           rows = not_miss_rows, cols = n_col, stack = TRUE)
  
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(E',n_row,'=0,"",IF(AND(E',n_row,'<>100,E',n_row,'<>1),"Vsota ni 100%",""))'),
               startCol = n_col+1, startRow = n_row)

  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col+1, rows = n_row,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFFF65"))
  
  setColWidths(wb = wb, sheet = one_dim_raking_var, cols = 1:n_col, widths = "auto")
  
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = 2, rows = not_miss_rows,
                        type = "between", rule = c(1,4),
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = 2, rows = not_miss_rows,
                        type = "between", rule = c(5,9),
                        style = createStyle(bgFill = "#FFECAF"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression", rule = 'AND(D2>0,D2<1)',
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression", rule = 'AND(D2>=1,D2<5)',
                        style = createStyle(bgFill = "#FFECAF"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = 'OR(COUNTIF(F2,"*_pop0"),COUNTIF(F2,"*_vzorec0"))',
                        style = createStyle(bgFill = "#FF4B4B"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = 'COUNTIF(F2,"*_ok")',
                        style = createStyle(bgFill = "#92D050"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('$E$',n_row,'=0'),
                        style = createStyle(bgFill = "#FFFFFF"))
  
  # Diagnostic text - critical errors
  writeFormula(wb = wb, sheet = one_dim_raking_var,
               x = 'IF(OR(H4<>"",H5<>""),"Kritična opozorila (procedura ne bo delovala):","")',
               xy = c(n_col+3, 3))
  
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(COUNTIF(F2:F',n_row-1,',"*_pop0"),"Ničelna populacijska in neničelna vzorčna margina","")'),
               startCol = n_col+3, startRow = 4)
  
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(COUNTIF(F2:F',n_row-1,',"*_vzorec0"),"Neničelna populacijska in ničelna vzorčna margina","")'),
               startCol = n_col+3, startRow = 5)
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col+3, rows = 4:5,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FF4B4B"))
  
  # Diagnostic text - recommendations
  writeFormula(wb = wb, sheet = one_dim_raking_var,
               x = 'IF(OR(H8<>"",H9<>""),"Priporočila (združevanje celic):","")',
               xy = c(n_col+3, 7))
  
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(AND(COUNTIF(F2:F',n_row-1,',"vzorec1-4_*"),COUNTIF(F2:F',n_row-1,',"*_vzorec1-4")),"Kritično majhne celice (n < 5) in populacijske margine (< 1%)",IF(COUNTIF(F2:F',n_row-1,',"*_vzorec1-4"),"Kritično majhne populacijske margine (< 1%)",IF(COUNTIF(F2:F',n_row-1,',"vzorec1-4_*"),"Kritično majhne celice (n < 5)","")))'),
               startCol = n_col+3, startRow = 8)
  
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(AND(COUNTIF(F2:F',n_row-1,',"vzorec5-9_*"),COUNTIF(F2:F',n_row-1,',"*_vzorec5-9")),"Majhne celice (n < 10) in populacijske margine (< 5%)",IF(COUNTIF(F2:F',n_row-1,',"*_vzorec5-9"),"Majhne populacijske margine (< 5%)",IF(COUNTIF(F2:F',n_row-1,',"vzorec5-9_*"),"Majhne celice (n < 10)","")))'),
               startCol = n_col+3, startRow = 9)
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col+3, rows = 8,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col+3, rows = 9,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFECAF"))
  
  addStyle(wb = wb, sheet = one_dim_raking_var, 
           style = createStyle(textDecoration = "bold"), 
           rows = c(3,7), cols = n_col+3, stack = TRUE)
  
  # Diagnostic text - ok
  writeFormula(wb = wb,
               sheet = one_dim_raking_var,
               x = paste0('IF(AND(COUNTIF(F2:F',n_row-1,',"*_ok"),$E$',n_row,'<>0),"Ustrezen vnos","")'),
               startCol = n_col+3, startRow = 1)
  
  conditionalFormatting(wb = wb, sheet = one_dim_raking_var,
                        cols = n_col+3, rows = 1,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#92D050"))
  
  setColWidths(wb = wb, sheet = one_dim_raking_var, cols = c(2, (n_col+1):(n_col+3)), widths = c(10, 12, 4, 48))
  
  protectWorksheet(wb = wb, sheet = one_dim_raking_var, protect = TRUE, 
                   lockFormattingColumns = FALSE, lockFormattingCells = FALSE)
  
  addStyle(wb = wb, sheet = one_dim_raking_var, 
           style = createStyle(numFmt = "NUMBER"), 
           rows = 2:n_row, cols = n_col-2, stack = TRUE)
  
  addStyle(wb = wb, sheet = one_dim_raking_var, 
           style = createStyle(numFmt = "NUMBER"), 
           rows = 2:n_row, cols = n_col-1, stack = TRUE)
  
  dataValidation(wb = wb, one_dim_raking_var, cols = n_col, rows = 2:(n_row-1),
                 type = "decimal", operator = "between", value = c(0, 100),
                 showInputMsg = FALSE, showErrorMsg = TRUE, allowBlank = FALSE)
}


# create two-dimensional contingency table to excel to input population margins
two_dim_excel <- function(orig_data, two_dim_raking_vars, wb, drop_zero, drop_zero_missing){
  
  two_dim_names <- paste0(two_dim_raking_vars, collapse = " x ")
  two_dim_names <- substr(two_dim_names, 1, 31) # Max length is limited to 31 characters for Excel sheet name
  temp_var1 <- two_dim_raking_vars[[1]]
  temp_var2 <- two_dim_raking_vars[[2]]
  
  temp_df <- data.frame(table(clean_data(orig_data[[temp_var1]]),
                              clean_data(orig_data[[temp_var2]]), useNA = "no"))
  temp_df$prop <- (temp_df$Freq/sum(temp_df$Freq))*100
  temp_df$popul <- rep(0, nrow(temp_df))
  temp_df$input <- rep(0, nrow(temp_df))
  temp_df <- rbind(temp_df, data.frame(Var1 = "Skupaj", 
                                       Var2 = "Skupaj",
                                       t(colSums(temp_df[,-c(1:2), drop = FALSE]))))
  colnames(temp_df) <- c(temp_var1, temp_var2, "Frekvenca", "Vzorcni %", "Populacijski (ciljni) %", "Vnos populacijskih margin (%)")
  
  if(drop_zero == TRUE){ # drop empty content categories (but leave empty missing categories)
    temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df[[temp_var1]] == "Missing" | temp_df[[temp_var2]] == "Missing"), ,drop = FALSE]
  }
  
  if(drop_zero_missing == TRUE){ # drop empty missing categories
    temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df$Frekvenca == 0 & temp_df[[temp_var1]] != "Missing" & temp_df[[temp_var2]] != "Missing"), ,drop = FALSE]
  }
  
  n_row <- nrow(temp_df) + 1
  n_col <- ncol(temp_df)
  
  addWorksheet(wb = wb, sheetName = two_dim_names, gridLines = FALSE)
  writeData(wb = wb,
            sheet = two_dim_names,
            x = temp_df,
            colNames = TRUE, rowNames = FALSE,
            borders = "all",
            headerStyle = createStyle(halign = "center", textDecoration = "bold", border = "TopBottomLeftRight"))
  addStyle(wb = wb,
           sheet = two_dim_names,
           rows = 1, cols = n_col,
           style = createStyle(fontColour = "#FFFFFF", fgFill = "#000000"),
           stack = TRUE)
  
  vec_row <- 2:(n_row-1) # vektor vrstic kot v excelu
  
  vec_miss <- temp_df[[1]][-(n_row-1)] == "Missing" | temp_df[[2]][-(n_row-1)] == "Missing" # vektor kjer so Missing polja
  
  miss_formula <- paste0("E",vec_row[vec_miss], collapse = ",")
  
  else_formula <- paste0("F",vec_row[!vec_miss], collapse = ",")
  
  not_miss_rows <- vec_row[!vec_miss]
  
  borders_row <- which(duplicated(temp_df[[2]], fromLast = TRUE) == FALSE)+1
  
  for(k in 2:n_row){
    if(temp_df[[1]][k-1] == "Missing" || temp_df[[2]][k-1] == "Missing"){ # Missing categories
      writeFormula(wb = wb, 
                   sheet = two_dim_names, 
                   x = paste0("IF(AND(D",k,"=0,F",k,"=0),0,IF(AND(D",k,"<>0,F",k,"=0),D",k,",IF(AND(D",k,"=0,F",k,"<>0),D",k,",IF(AND(D",k,"<>0,F",k,"<>0,F",k,">D",k,"),D",k,",IF(AND(D",k,"<>0,F",k,"<>0,F",k,"<=D",k,"),F",k,")))))"), 
                   startCol = n_col-1, startRow = k)
      
      addStyle(wb = wb, sheet = two_dim_names,
               style = createStyle(fgFill = "#D9D9D9"),
               rows = k, cols = 1:n_col, stack = TRUE)
      
      addStyle(wb = wb, two_dim_names, 
               style = createStyle(locked = FALSE,
                                   textDecoration = "bold"), 
               rows = k, cols = n_col, stack = TRUE)
      
    } else if(temp_df[[1]][k-1] == "Skupaj" && temp_df[[2]][k-1] == "Skupaj"){ # Sum category
      writeFormula(wb = wb, 
                   sheet = two_dim_names, 
                   x = paste0("SUM(E2:E",n_row-1,")"), 
                   startCol = n_col-1, startRow = n_row)
      
      writeFormula(wb = wb, 
                   sheet = two_dim_names, 
                   x = paste0("SUM(F2:F",n_row-1,")"), 
                   startCol = n_col, startRow = n_row)
      
      addStyle(wb = wb, 
               sheet = two_dim_names, 
               style = createStyle(border = "top", 
                                   borderStyle = "double",
                                   textDecoration = "bold"), 
               rows = k, cols = 1:n_col, stack = TRUE)
      
    } else{ # Content categories
      if(any(vec_miss)){
        writeFormula(wb = wb, 
                     sheet = two_dim_names, 
                     x = paste0("IFERROR(F",k,"*(100-SUM(",miss_formula,"))/(SUM(",else_formula,")),0)"), 
                     startCol = n_col-1, startRow = k)
      } else{
        writeFormula(wb = wb, 
                     sheet = two_dim_names, 
                     x = paste0("IFERROR((F",k,"/SUM(",else_formula,"))*100,0)"), 
                     startCol = n_col-1, startRow = k)
      }
      
      writeFormula(wb = wb, 
                   sheet = two_dim_names,
                   x = paste0('IF(AND(C',k,'>=1,C',k,'<=4),"vzorec1-4",IF(AND(C',k,'>=5,C',k,'<=9),"vzorec5-9","ok"))&"_"&IF(AND(E',k,'=0,C',k,'>0,$F$',n_row,'<>0),"pop0",IF(AND(C',k,'=0,E',k,'>0,$F$',n_row,'<>0),"vzorec0",IF(AND(E',k,'>0,E',k,'<1),"vzorec1-4",IF(AND(E',k,'>=1,E',k,'<5),"vzorec5-9","ok"))))'),
                   startCol = n_col+1, startRow = k)
    }
  }
  
  addStyle(wb = wb, 
           sheet = two_dim_names, 
           style = createStyle(fontColour = "#FFFFFF"), 
           rows = not_miss_rows, cols = n_col+1, stack = TRUE)
  
  addStyle(wb = wb,
           sheet = two_dim_names,
           style = createStyle(border = "bottom", borderStyle = "double"),
           rows = borders_row, cols = 1:n_col, stack = TRUE, gridExpand = TRUE)
  
  addStyle(wb = wb, 
           sheet = two_dim_names, 
           style = createStyle(locked = FALSE,
                               textDecoration = "bold"), 
           rows = not_miss_rows, cols = n_col, stack = TRUE)
  
  protectWorksheet(wb = wb,
                   sheet = two_dim_names, 
                   protect = TRUE, 
                   lockFormattingColumns = FALSE, lockFormattingCells = FALSE)
  
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(F',n_row,'=0,"",IF(AND(F',n_row,'<>100,F',n_row,'<>1),"Vsota ni 100%",""))'),
               startCol = n_col+1, startRow = n_row)

  conditionalFormatting(wb = wb,
                        sheet = two_dim_names,
                        cols = n_col+1, rows = n_row,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFFF65"))
  
  setColWidths(wb = wb, sheet = two_dim_names, cols = 1:n_col, widths = "auto")
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = 3, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('AND(C',not_miss_rows[1],'>0,C',not_miss_rows[1],'<=4,G',not_miss_rows[1],'<>"")'),
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = 3, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('AND(C',not_miss_rows[1],'>=5,C',not_miss_rows[1],'<=9,G',not_miss_rows[1],'<>"")'),
                        style = createStyle(bgFill = "#FFECAF"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('AND(E',not_miss_rows[1],'>0,E',not_miss_rows[1],'<1,G',not_miss_rows[1],'<>"")'),
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('AND(E',not_miss_rows[1],'>=1,E',not_miss_rows[1],'<5,G',not_miss_rows[1],'<>"")'),
                        style = createStyle(bgFill = "#FFECAF"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('OR(COUNTIF(G',not_miss_rows[1],',"*_pop0"),COUNTIF(G',not_miss_rows[1],',"*_vzorec0"))'),
                        style = createStyle(bgFill = "#FF4B4B"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('COUNTIF(G',not_miss_rows[1],',"*_ok")'),
                        style = createStyle(bgFill = "#92D050"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col-1, rows = not_miss_rows,
                        type = "expression",
                        rule = paste0('AND($F$',n_row,'=0,COUNTIF(G',not_miss_rows[1],',"*_ok"))'),
                        style = createStyle(bgFill = "#FFFFFF"))
  
  # Diagnostic text - critical errors
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = 'IF(OR(I4<>"",I5<>""),"Kritična opozorila (procedura ne bo delovala):","")',
               xy = c(n_col+3, 3))
  
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(COUNTIF(G2:G',n_row-1,',"*pop0"),"Ničelna populacijska in neničelna vzorčna margina","")'),
               startCol = n_col+3, startRow = 4)
  
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(COUNTIF(G2:G',n_row-1,',"*vzorec0"),"Neničelna populacijska in ničelna vzorčna margina","")'),
               startCol = n_col+3, startRow = 5)
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col+3, rows = 4:5,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FF4B4B"))
  
  # Diagnostic text - recommendations
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = 'IF(OR(I8<>"",I9<>""),"Priporočila (združevanje celic):","")',
               xy = c(n_col+3, 7))
  
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(AND(COUNTIF(G2:G',n_row-1,',"vzorec1-4_*"),COUNTIF(G2:G',n_row-1,',"*_vzorec1-4")),"Kritično majhne celice (n < 5) in populacijske margine (< 1%)",IF(COUNTIF(G2:G',n_row-1,',"*_vzorec1-4"),"Kritično majhne populacijske margine (< 1%)",IF(COUNTIF(G2:G',n_row-1,',"vzorec1-4_*"),"Kritično majhne celice (n < 5)","")))'),
               startCol = n_col+3, startRow = 8)
  
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(AND(COUNTIF(G2:G',n_row-1,',"vzorec5-9_*"),COUNTIF(G2:G',n_row-1,',"*_vzorec5-9")),"Majhne celice (n < 10) in populacijske margine (< 5%)",IF(COUNTIF(G2:G',n_row-1,',"*_vzorec5-9"),"Majhne populacijske margine (< 5%)",IF(COUNTIF(G2:G',n_row-1,',"vzorec5-9_*"),"Majhne celice (n < 10)","")))'),
               startCol = n_col+3, startRow = 9)
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col+3, rows = 8,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFC000"))
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col+3, rows = 9,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#FFECAF"))
  
  addStyle(wb = wb, sheet = two_dim_names, 
           style = createStyle(textDecoration = "bold"), 
           rows = c(3,7), cols = n_col+3, stack = TRUE)
  
  # Diagnostic text - ok
  writeFormula(wb = wb,
               sheet = two_dim_names,
               x = paste0('IF(AND(COUNTIF(G2:G',n_row-1,',"*_ok"),$F$',n_row,'<>0),"Ustrezen vnos","")'),
               startCol = n_col+3, startRow = 1)
  
  conditionalFormatting(wb = wb, sheet = two_dim_names,
                        cols = n_col+3, rows = 1,
                        type = "expression", rule = '<>""',
                        style = createStyle(bgFill = "#92D050"))
  
  setColWidths(wb = wb, sheet = two_dim_names, cols = c(3, (n_col+1):(n_col+3)), widths = c(10, 12, 4, 48))
  
  addStyle(wb = wb, sheet = two_dim_names, 
           style = createStyle(numFmt = "NUMBER"), 
           rows = 2:n_row, cols = n_col-2, stack = TRUE)
  
  addStyle(wb = wb, sheet = two_dim_names, 
           style = createStyle(numFmt = "NUMBER"), 
           rows = 2:n_row, cols = n_col-1, stack = TRUE)
  
  dataValidation(wb = wb, sheet = two_dim_names, cols = n_col, rows = 2:(n_row-1),
                 type = "decimal", operator = "between", value = c(0, 100),
                 showInputMsg = FALSE, showErrorMsg = TRUE, allowBlank = FALSE)
}


# export tables to input margins to excel file
input_file_output_tables_excel <- function(orig_data = NULL, 
                                           one_dimensional_raking_variables = NULL, # accept vector
                                           two_dimensional_raking_variables = NULL, # accept list only
                                           file_name = NULL,
                                           drop_zero,
                                           drop_zero_missing){
  
  all_raking_variables_list <- c(as.list(one_dimensional_raking_variables),
                                 two_dimensional_raking_variables)
  
  workbook <- createWorkbook()
  
  for(i in seq_along(all_raking_variables_list)){
    
    if(length(all_raking_variables_list[[i]]) == 1){
      
      one_dim_excel(orig_data = orig_data, 
                    one_dim_raking_var = all_raking_variables_list[[i]], 
                    wb = workbook,
                    drop_zero = drop_zero,
                    drop_zero_missing = drop_zero_missing)
      
    } else if(length(all_raking_variables_list[[i]]) == 2){
      
      two_dim_excel(orig_data = orig_data, 
                    two_dim_raking_vars = all_raking_variables_list[[i]],
                    wb = workbook, 
                    drop_zero = drop_zero,
                    drop_zero_missing = drop_zero_missing)
    }
  }
  
  protectWorkbook(wb = workbook, lockStructure = TRUE)
  
  saveWorkbook(wb = workbook, file = file_name, overwrite = TRUE)
}