This repository contains an R script to semi-automatically code answers to the question "*What do you think are the most important problems in our country, today?*" into one of several pre-defined categories. The script compares each answer against a set of fixed keywords and pre-coded answers from earlier waves using exact and fuzzy matching. If SPSS is one's primary tool and gets preference over R, there's a function to generate matching SPSS recode syntax.


### 1. Code answers
```R
[...]
```

### 2. Create SPSS syntax
Load functions `xlsx` and `xlsx_to_sps`:
```R
install.packages("xlsx")
library(xlsx)

xlsx_to_sps <- eval(parse(text = source("https://raw.githubusercontent.com/siardv/MIP/main/xlsx_to_sps.R")[1]))
```
| Arguments    	|                                                                             	|
|--------------	|-----------------------------------------------------------------------------	|
| xlsx        	| Directory to Excel file(s) containing pre-categorized answers               	|
| id          	| Individual-level identifier                                                 	|
| answer      	| Original answer to open-ended question                                      	|
| value_labels 	| List with categories/labels used in the Excel file and corresponding values 	|
| user_na     	| Which values are missing values?                                            	|
| save_to     	| Directory to save the .sps file                                             	|

Example:
```R
xlsx_to_sps(path = ".../Documents/MIP/mip.xlsx", 
            id = 1,
            answer = 2,
            value_labels = list(value = c(1:24, 991, 992, 999),
                                label = c("Economy / Financial situation", "Social security", "Politics", "Crime", "Defense",
                                          "Healthcare", "Education", "Income / Prince levels / Taxes", "Employment",
                                          "Traffic / Mobility", "Housing", "Environment", "Population", "Minorities", 
                                          "Norms and values",  "Media", "European integration", "Inequality / Poverty", 
                                          "Intolerance / Discrimination", "Foreign policy / International security", 
                                          "Regulation / Big government", "Polarisation / Dividedness", "Immigration", 
                                          "Corona", "There are no problems", "No other problems", "DK/NA/Cannot be coded")),
            user_na = c(991, 992, 999),
            save_to = ".../Desktop")
               
# or load multiple files:
path = paste0(".../Documents/MIP/", c("mip_1", "mip_2", "mip_3"), ".xlsx")
```
