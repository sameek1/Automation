## library
  library(readxl)
  library(readr)
  library(dplyr)
  library(janitor)  
  library(reshape2)
  library(purrr)
  library(openxlsx)
  #library(rpivotTable)
  #library(XLConnect)
  
## file mappings
  df_read_settings <- read_excel("file_read_mapping_summary.xlsx",sheet = "read_settings")
  df_col_mapping <- read_excel("file_read_mapping_summary.xlsx",sheet = "mapping2")
  
## file mapping for data type formatting  
  df_format <- read.xlsx("file_read_mapping_summary.xlsx", sheet = "Format", colNames = TRUE, startRow = 1 )
  
  data_directory <- "./data4/"

## create function to read data
  data_reader <- function(result){
    ## check for files with sheet name
    df1 <- df_read_settings %>% filter(output == result)
    print(df1)
    output <- data.frame()
    for (i in 1:nrow(df1)) {
      data_path <- paste0(data_directory,df1$dir[i])
      print(data_path)
      files <- dir(data_path, pattern = "*.xlsx")
      for (j in 1:length(files)){
        if((TRUE %in% (df1$sheet_name[i] %in% excel_sheets(file.path(data_path, files)[j])))){
          # print(df1$sheet_name[i])
          tmp <- read_excel(file.path(data_path,files)[j],sheet = df1$sheet_name[i],skip = df1$skip[i])
          ncol.temp <- ncol(tmp)
          tmp2 <-  read_excel(file.path(data_path,files)[j],sheet = df1$sheet_name[i],col_types = rep("text", ncol.temp),skip = df1$skip[i])
          tmp2$file_source <- file.path(data_path,files)[j]
          tmp2$sheet_name <- df1$sheet_name[i]
          print(tmp2$sheet_name)
          print(tmp2$file_source)
          ## rename col
          df_names <- df_col_mapping %>% filter(directory == df1$dir[i],sheet_name == df1$sheet_name[i])
          if(nrow(df_names)>0){
            existing <- match(df_names$old_col,names(tmp2))
            names(tmp2)[na.omit(existing)] <- df_names$new_col[which(!is.na(existing))]
          }
          
          
         
          
          
          
          
          
          

          ## append data
          if("file_source" %in% names(output)){
            output <- full_join(output,tmp2)
          }else{
            output <- tmp2
          }
        }
      }
    }
    return(output)
  }
  

  
  #<--------- Use the below code for consolidation to overall summary file ------->
  df_overall <- data_reader("Overall")
  df1 <- df_overall[c("Revised Audience","Week of","USD Spend","Date (mm/dd/yyyy)","Market","Audience","Content Type","Creative USP Focus","Creative Lifestyle Message","Channel","Currency","Spend","Impressions","Clicks","Video Views","Video Completes",".com Land - Click Conversion",".com Land - View Conversion","Where to Buy - Click","Where to Buy - View","Buy Now - Click","Buy Now - View","Add to Cart - Click","Add to Cart - View","Order Confirmation - Click","Order Confirmation - View","Social Engagements","Likes","Shares","Comments","Reactions","Search Impressions","Keyword Rank","TV GRPs","OOH/Print Impacts","Conversion Tag name ( Column for data taken from APIs only)","Total Conversions ( Column for data taken from APIs only)","Market Category","Channel Type","Market Tier","Channel_New","Channel_New(2)")]
  
  
  #<--------- Use the code to consolidate to weekly consolidated file ------> 
  
  df_dsp <- data_reader("dsp")
  df_dsp <- df_dsp[c("Market","Audience","Content Type","Creative USP Focus","Creative Lifestyle Message","Date","Channel","Publisher/Site","Creative Type","Placement Name","Rate Type","Currency","Spend","Impressions","Clicks","Video Views","Video Completes (100%)",".com Land - Click Conversion",".com Land - View Conversion","Where to Buy - Click","Where to Buy - View","Buy Now - Click","Buy Now - View","Add to Cart - Click","Add to Cart - View","Order Confirmation - Click","Order Confirmation - View","Star Explore Page - Click","Star Explore Page - View","Conversion Tag Name","Total Conversions")]
  df_social <- data_reader("social")
  df_social <- df_social[c("Market","Audience","Content Type","Creative USP Focus","Creative Lifestyle Message","Date","Channel","Publisher/Site","Creative Type","Placement Name","Rate Type","Currency","Spend","Impressions","Clicks","Video Views","Video Completes (100%)","Likes","Shares","Comments",".com Land - Click Conversion",".com Land - View Conversion","Where to Buy - Click","Where to Buy - View","Buy Now - Click","Buy Now - View","Add to Cart - Click","Add to Cart - View","Order Confirmation - Click","Order Confirmation - View","Star Explore Page - Click","Star Explore Page - View","Conversion Tag Name","Total Conversions","Post Type")]
  df_search <- data_reader("search")
  df_search <- df_search[c("Market","Date","Channel","Publisher/Site","Keyword Group","Rate Type","Currency","Spend","Impressions","Clicks","SOV","Avg Ranking","Quality Score",".com Land - Click Conversion",".com Land - View Conversion","Where to Buy - Click","Where to Buy - View","Buy Now - Click","Buy Now - View","Add to Cart - Click","Add to Cart - View","Order Confirmation - Click","Order Confirmation - View","Star Explore Page - Click","Star Explore Page - View","Conversion Tag Name","Total Conversions")]
  df_offline <- data_reader("Offline")
  options(max.print = 100000000)
  
  #all the variables (metrics) which need to be converted to Numeric datatype in Excel 
  #listA <- c("OOH/Print Impacts","TV GRPs","SOV","Keyword Rank","Quality Score","Likes","Shares","Comments","Total Conversions","Impressions","Clicks","Spend","Video Views","Video Completes (100%)",".com Land - Click Conversion",".com Land - View Conversion","Where to Buy - Click","Where to Buy - View","Buy Now - Click","Buy Now - View","Add to Cart - Click","Add to Cart - View","Order Confirmation - Click","Order Confirmation - View","Star Explore Page - Click","Star Explore Page - View","Star Shop Page - Click","Star Shop Page - View","Star Amazon product Pages - Click","Star Amazon product Pages - View")
  
  
  # <------------------- formatting to appropriate data types --------------------->
  
  isNumeric <- as.list(df_format$isNumeric)
  isNumeric <- isNumeric[!is.na(isNumeric)]
  isDate <- (as.list(df_format$`isDate`))
  isDate <- isDate[!is.na(isDate)]
  
  #converting the isNumeric variables in the list to numeric datatype  
  for(i in isNumeric){
    print(i)
    if(i %in% colnames(df_dsp)){
    df_dsp[[i]] <- as.numeric(as.character(df_dsp[[i]]))
    }
    
    if(i %in% colnames(df_social)){
      df_social[[i]] <- as.numeric(as.character(df_social[[i]]))
    }
    
    if(i %in% colnames(df_search)){
      df_search[[i]] <- as.numeric(as.character(df_search[[i]]))
    }
    
    if(i %in% colnames(df_offline)){
      df_offline[[i]] <- as.numeric(as.character(df_offline[[i]]))
    }
    
  }
  
  
  #<------ for overall consolidation --------->
  for(i in isNumeric){
    if(i %in% colnames(df1)){
      df1[[i]] <- as.numeric(as.character(df1[[i]]))
    }
  }

    
  #converting the isNumeric variables in the list to numeric datatype 
  df_dsp$`Date` <- excel_numeric_to_date(as.numeric(df_dsp$`Date`))
  df_social$`Date` <- excel_numeric_to_date(as.numeric(df_social$`Date`))
  df_search$`Date` <- excel_numeric_to_date(as.numeric(df_search$`Date`))
  df_offline$`Date` <- excel_numeric_to_date(as.numeric(df_offline$`Date`))
  
  df1$`Week of` <- excel_numeric_to_date(as.numeric(df1$`Week of`))
  df1$`Date (mm/dd/yyyy)` <- excel_numeric_to_date(as.numeric(df1$`Date (mm/dd/yyyy)`))

  # <--------- Use for openXlsx only (outputs Excel file) ----------->
  wb1 <- createWorkbook()
  addWorksheet(wb1,"DigitalDisplay Weekly Reporting")
  addWorksheet(wb1,"Social Weekly Reporting")
  addWorksheet(wb1,"Search Weekly Reporting")
  addWorksheet(wb1,"Offline")
  writeData(wb1,"DigitalDisplay Weekly Reporting",df_dsp)
  writeData(wb1,"Social Weekly Reporting",df_social)
  writeData(wb1,"Search Weekly Reporting",df_search)
  writeData(wb1,"Offline",df_offline)
  saveWorkbook(wb1,"Summary_June_10_v2.xlsx",overwrite = TRUE)
  
  
  # loop and write <----- Use for outputting in CSV only ------->
  for (i in 1:length(unique(df_read_settings$output))) {
    tmp <- data_reader(unique(df_read_settings$output)[i])
    file_name <- paste0(unique(df_read_settings$output)[i],".csv")
    write_csv(tmp,file_name,na="")
  }
  
   
 # <---------- Use for XLConnect only  ---------->
  # wb <- loadWorkbook("test2.xlsx",create=TRUE)
  # createSheet(wb,name="dsp")
  # createSheet(wb,name="social")
  # createSheet(wb,name="search")
  # writeWorksheet(wb,df_dsp,"dsp")
  # writsaveWorkbook(wb)
  # writeWorksheet(wb,df_social,"social")
  # saveWorkbook(wb)
  # writeWorksheet(wb,df_search,"search")
  # saveWorkbook(wb)
  
  
 # <------- Use for csv output --------->
  
  
  
  
## read in the mapping file 
  
  