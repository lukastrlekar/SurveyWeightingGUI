# DATA READING AND MANIPULATION ------------------------------------------------------------

# Read original raw data
load_file <- function(path, ext, decimal_sep, delim, na_strings) {
  switch(ext,
         # sav = foreign::read.spss(file = path,
         #                          to.data.frame = TRUE,
         #                          use.missings = TRUE,
         #                          use.value.labels = TRUE),
         sav = haven::read_sav(file = path, user_na = TRUE),
         rds = readRDS(file = path),
         csv = read.csv(file = path, header = TRUE, sep = delim, dec = decimal_sep, na.strings = trimws(strsplit(x = na_strings, split = ",")[[1]])),
         txt = read.delim(file = path, header = TRUE, sep = delim, dec = decimal_sep, na.strings = trimws(strsplit(x = na_strings, split = ",")[[1]])))
}

# Load the Rdata into a new environment to avoid side effects
load_to_environment <- function(RData, env = new.env()) {
  load(RData, env)
  return(env[[names(env)[1]]])
}

# Very quick one-dimensional weighted table function with no external dependencies
weighted_table = function(x, weights) {
  vapply(split(weights, x), sum, 0, na.rm = TRUE)
}

# internal function for some data manipulation
clean_data <- function(variable = NULL) {
  # first convert user-defined missing values to R's NAs (in case of SPSS uploaded file, otherwise everything stays the same)
  # then convert to factor
  temp_fac <- haven::as_factor(labelled::user_na_to_na(variable))
  
  # transform NAs to new level "Missing", else just add new empty level "Missing"
  temp_fac <- addNA(temp_fac, ifany = FALSE)
  levels(temp_fac)[is.na(levels(temp_fac))] <- "Missing"
  
  return(temp_fac)
}

# display tables in input margins excel tab
display_tables <- function(orig_data = NULL,
                           one_dimensional_raking_variables = NULL,
                           two_dimensional_raking_variables = NULL,
                           drop_zero,
                           drop_zero_missing){
  
  all_raking_variables_list <- c(as.list(one_dimensional_raking_variables),
                                 two_dimensional_raking_variables)
  
  tables <- vector(mode = "list", length = length(all_raking_variables_list))
  
  for(i in seq_along(all_raking_variables_list)){
    
    if(length(all_raking_variables_list[[i]]) == 1){
      temp_var <- all_raking_variables_list[[i]]
      
      temp_df <- data.frame(table(clean_data(orig_data[[temp_var]]), useNA = "no"))
      temp_df$prop <- (temp_df$Freq/sum(temp_df$Freq))*100
      temp_df <- rbind(temp_df, data.frame(Var1 = "Skupaj", t(colSums(temp_df[,-1, drop = FALSE]))))
      colnames(temp_df) <- c(temp_var, "Frekvenca", "Odstotek (%)")
      
      if(drop_zero == TRUE){ # drop empty content categories
        temp_df <- temp_df[temp_df$Frekvenca > 0 | temp_df[[temp_var]] == "Missing", ,drop = FALSE]
      }
      
      if(drop_zero_missing == TRUE){ # drop emtpy missing categories
        temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df$Frekvenca == 0 & temp_df[[temp_var]] != "Missing"), ,drop = FALSE]
      }
      
      tables[[i]] <- temp_df
      names(tables)[[i]] <- temp_var
      
    } else if(length(all_raking_variables_list[[i]]) == 2){
      temp_var1 <- all_raking_variables_list[[i]][[1]]
      temp_var2 <- all_raking_variables_list[[i]][[2]]
      
      temp_df <- data.frame(table(clean_data(orig_data[[temp_var1]]),
                                  clean_data(orig_data[[temp_var2]]), useNA = "no"))
      temp_df$prop <- (temp_df$Freq/sum(temp_df$Freq))*100
      temp_df <- rbind(temp_df, data.frame(Var1 = "Skupaj", 
                                           Var2 = "Skupaj",
                                           t(colSums(temp_df[,-c(1:2), drop = FALSE]))))
      colnames(temp_df) <- c(temp_var1, temp_var2, "Frekvenca", "Odstotek (%)")
      
      if(drop_zero == TRUE){ # drop empty content categories (but leave empty missing categories)
        temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df[[temp_var1]] == "Missing" | temp_df[[temp_var2]] == "Missing"), ,drop = FALSE]
      }
      
      if(drop_zero_missing == TRUE){ # drop empty missing categories
        temp_df <- temp_df[temp_df$Frekvenca > 0 | (temp_df$Frekvenca == 0 & temp_df[[temp_var1]] != "Missing" & temp_df[[temp_var2]] != "Missing"), ,drop = FALSE]
      }
      
      tables[[i]] <- temp_df
      names(tables)[[i]] <- paste0(temp_var1," x ", temp_var2)
    }
  }
  
  return(tables[lengths(tables) != 0])
}

# tables to input margins in input margins app tab
input_tables <- function(orig_data = NULL,
                         one_dimensional_raking_variables = NULL,
                         two_dimensional_raking_variables = NULL,
                         drop_zero,
                         drop_zero_missing){
  
  all_raking_variables_list <- c(as.list(one_dimensional_raking_variables),
                                 two_dimensional_raking_variables)
  
  tables <- vector(mode = "list", length = length(all_raking_variables_list))
  
  for(i in seq_along(all_raking_variables_list)){
    
    if(length(all_raking_variables_list[[i]]) == 1){
      temp_var <- all_raking_variables_list[[i]]
      
      temp_df <- as.data.frame(table(clean_data(orig_data[[temp_var]]), useNA = "no"))
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
      
      tables[[i]] <- temp_df
      names(tables)[[i]] <- temp_var
      
    } else if(length(all_raking_variables_list[[i]]) == 2){
      temp_var1 <- all_raking_variables_list[[i]][[1]]
      temp_var2 <- all_raking_variables_list[[i]][[2]]
      
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
      
      tables[[i]] <- temp_df
      names(tables)[[i]] <- paste0(temp_var1," x ", temp_var2)
    }
  }
  
  return(tables[lengths(tables) != 0])
}