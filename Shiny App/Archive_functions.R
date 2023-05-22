
# Archive functions (not used) --------------------------------------------



# clean_data <- function(orig_data = NULL, 
#                        one_dimensional_raking_variables = NULL,
#                        two_dimensional_raking_variables = NULL) {
#   
#   all_raking_variables <- unlist(c(one_dimensional_raking_variables,
#                                    two_dimensional_raking_variables))
#   
#   selected_data <- orig_data[,all_raking_variables, drop = FALSE]
#   
#   # convert every selected variable to a factor (just in case)
#   selected_data[] <- lapply(selected_data, as.factor) 
#   
#   for(i in seq_along(all_raking_variables)){
#     # convert NA as new level "Missing"  
#     if(any(is.na(selected_data[[all_raking_variables[i]]]))){
#       # selected_data[[all_raking_variables[i]]] <- forcats::fct_explicit_na(selected_data[[all_raking_variables[i]]],
#       #                                                                      na_level = "Missing")
#       
#       selected_data[[all_raking_variables[i]]] <- addNA(selected_data[[all_raking_variables[i]]])
#       levels(selected_data[[all_raking_variables[i]]])[is.na(levels(selected_data[[all_raking_variables[i]]]))] <- "Missing"
#       
#     } else{
#       # selected_data[[all_raking_variables[i]]] <- forcats::fct_expand(selected_data[[all_raking_variables[i]]], "Missing")
#       levels(selected_data[[all_raking_variables[i]]]) <- c(levels(selected_data[[all_raking_variables[i]]]), "Missing") 
#     }
#   }
#   
#   return(selected_data)
# }

# function for displaying preview tables
# display_tables <- function(orig_data = NULL,
#                            one_dimensional_raking_variables = NULL,
#                            two_dimensional_raking_variables = NULL,
#                            drop_zero,
#                            drop_zero_missing){
#   
#   selected_data <- clean_data(orig_data = orig_data, 
#                               one_dimensional_raking_variables = one_dimensional_raking_variables,
#                               two_dimensional_raking_variables = two_dimensional_raking_variables)
#   
#   all_raking_variables_list <- c(as.list(one_dimensional_raking_variables),
#                                  two_dimensional_raking_variables)
#   
#   tables <- vector(mode = "list", length = length(all_raking_variables_list))
#   
#   for(i in seq_along(all_raking_variables_list)){
#     
#     if(length(all_raking_variables_list[[i]]) == 1){
#       temp_df <- 
#         count(x = selected_data, .data[[all_raking_variables_list[[i]]]], 
#               name = "Frekvenca",
#               .drop = FALSE) %>% 
#         mutate(`Odstotek (%)` = (Frekvenca/sum(Frekvenca))*100) %>%   
#         bind_rows(summarise(.,
#                             across(where(is.numeric), sum),
#                             across(where(is.factor), ~ "Skupaj"))) %>% 
#         as.data.frame()
#       
#       if(drop_zero == TRUE){ # drop empty content categories
#         temp_df <- temp_df %>% 
#           filter(Frekvenca > 0 | .data[[all_raking_variables_list[[i]]]] == "Missing")
#       }
#       
#       if(drop_zero_missing == TRUE){ # drop emtpy missing categories
#         temp_df <- temp_df %>% 
#           filter(Frekvenca > 0 | (Frekvenca == 0 & .data[[all_raking_variables_list[[i]]]] != "Missing"))
#       }
#       
#       tables[[i]] <- temp_df
#       names(tables)[[i]] <- all_raking_variables_list[[i]]
#       
#     } else if(length(all_raking_variables_list[[i]]) == 2){
#       temp_df <- 
#         count(x = selected_data, 
#               .data[[all_raking_variables_list[[i]][[1]]]],
#               .data[[all_raking_variables_list[[i]][[2]]]],
#               name = "Frekvenca",
#               .drop = FALSE) %>%   
#         mutate(`Odstotek (%)` = (Frekvenca/sum(Frekvenca))*100) %>% 
#         bind_rows(summarise(.,
#                             across(where(is.numeric), sum),
#                             across(where(is.factor), ~ "Skupaj"))) %>% 
#         as.data.frame()
#       
#       if(drop_zero == TRUE){ # drop empty categories (but leave empty Missing categories)
#         temp_df <- temp_df %>% 
#           filter(Frekvenca > 0 | (.data[[all_raking_variables_list[[i]][[1]]]] == "Missing" | .data[[all_raking_variables_list[[i]][[2]]]] == "Missing"))
#       }
#       
#       if(drop_zero_missing == TRUE){ # drop empty Missing categories
#         temp_df <- temp_df %>% 
#           filter(Frekvenca > 0 | (Frekvenca == 0 & .data[[all_raking_variables_list[[i]][[1]]]] != "Missing" & .data[[all_raking_variables_list[[i]][[2]]]] != "Missing"))
#       }
#       
#       tables[[i]] <- temp_df
#       names(tables)[[i]] <- paste0(all_raking_variables_list[[i]][[1]]," x ", all_raking_variables_list[[i]][[2]])
#     }
#   }
#   
#   return(tables[lengths(tables) != 0])
# }

# input_file_output_tables <- function(orig_data = NULL, 
#                                      one_dimensional_raking_variables = NULL, # accept vector
#                                      two_dimensional_raking_variables = NULL, # accept list only
#                                      file_name = NULL,
#                                      drop_zero,
#                                      drop_zero_missing){
#   
#   selected_data <- clean_data(orig_data = orig_data, 
#                               one_dimensional_raking_variables = one_dimensional_raking_variables,
#                               two_dimensional_raking_variables = two_dimensional_raking_variables)
#   
#   
#   all_raking_variables_list <- c(as.list(one_dimensional_raking_variables),
#                                  two_dimensional_raking_variables)
#   
#   workbook <- createWorkbook()
#   
#   for(i in seq_along(all_raking_variables_list)){
#     
#     if(length(all_raking_variables_list[[i]]) == 1){
#       
#       one_dim_excel(selected_data = selected_data, 
#                     one_dim_raking_var = all_raking_variables_list[[i]], 
#                     wb = workbook,
#                     drop_zero = drop_zero,
#                     drop_zero_missing = drop_zero_missing)
#       
#     } else if(length(all_raking_variables_list[[i]]) == 2){
#       
#       two_dim_excel(selected_data = selected_data, 
#                     two_dim_raking_vars = all_raking_variables_list[[i]],
#                     wb = workbook, 
#                     drop_zero = drop_zero,
#                     drop_zero_missing = drop_zero_missing)
#     }
#   }
#   
#   protectWorkbook(wb = workbook, lockStructure = TRUE)
#   
#   saveWorkbook(wb = workbook, file = file_name, overwrite = TRUE)
# }

# read_variables_excel <- function(sheet_list){
#   one_dimensional_raking_variables <- NULL
#   two_dimensional_raking_variables <- NULL
#   
#   for(i in seq_len(length(sheet_list))){
#     n_col <- ncol(sheet_list[[i]])
#     vec <- sapply(sheet_list[[i]][,-((n_col-1):n_col), drop = FALSE], is.character)
#     variables <- names(vec[vec==TRUE])
#     
#     if(length(variables) == 1){
#       one_dimensional_raking_variables <- c(one_dimensional_raking_variables, variables)
#     } else if(length(variables) == 2){
#       two_dimensional_raking_variables <- c(two_dimensional_raking_variables,
#                                             paste0(variables, collapse = " x "))
#     }
#   }
#   
#   return(list(one_dimensional_raking_variables,
#               two_dimensional_raking_variables))
# }


# excel

# create contingency table
# temp_df <- 
#   count(x = selected_data, .data[[one_dim_raking_var]], 
#         .drop = FALSE) %>% 
#   mutate(`Vzorcni %` = (n/sum(n))*100,
#          `Populacijski %` = 0,
#          `Vnos populacijskih margin (%)` = 0) %>%   
#   bind_rows(summarise(.,
#                       across(where(is.numeric), sum),
#                       across(where(is.factor), ~ "Sum"))) %>% 
#   as.data.frame()
# 
# if(drop_zero == TRUE){ # drop empty content categories
#   temp_df <- temp_df %>% 
#     filter(n > 0 | .data[[one_dim_raking_var]] == "Missing")
# }
# 
# if(drop_zero_missing == TRUE){ # drop emtpy missing categories
#   temp_df <- temp_df %>% 
#     filter(n > 0 | (n == 0 & .data[[one_dim_raking_var]] != "Missing"))
# }


# temp_df <- 
#   count(x = selected_data, 
#         .data[[two_dim_raking_vars[[1]]]],
#         .data[[two_dim_raking_vars[[2]]]],
#         .drop = FALSE) %>%   
#   mutate(`Vzorcni %` = (n/sum(n))*100,
#          `Populacijski %` = 0,
#          `Vnos populacijskih margin (%)` = 0) %>% 
#   bind_rows(summarise(.,
#                       across(where(is.numeric), sum),
#                       across(where(is.factor), ~ "Sum"))) %>% 
#   as.data.frame()
# 
# if(drop_zero == TRUE){ # drop empty categories (but leave empty Missing categories)
#   temp_df <- temp_df %>% 
#     filter(n > 0 | (.data[[two_dim_raking_vars[[1]]]] == "Missing" | .data[[two_dim_raking_vars[[2]]]] == "Missing"))
# }
# 
# if(drop_zero_missing == TRUE){ # drop empty Missing categories
#   temp_df <- temp_df %>% 
#     filter(n > 0 | (n == 0 & .data[[two_dim_raking_vars[[1]]]] != "Missing" & .data[[two_dim_raking_vars[[2]]]] != "Missing"))
# }


