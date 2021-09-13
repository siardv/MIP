#### Generate SPSS syntax from Excel files with pre-categorized answers to open-ended questions.

Load function from GitHub
```R
xlsx_to_syntax <- eval(parse(text = source("https://raw.githubusercontent.com/siardv/SKON/main/xlsx_to_spss.R")[1]))
```
| Arguments    	|                                                                             	|
|--------------	|-----------------------------------------------------------------------------	|
| xlsx        	| Directory to Excel file(s) containing pre-categorized answers               	|
| id          	| Individual-level identifier                                                 	|
| answer      	| Original answer to open-ended question                                      	|
| value_label 	| List with categories/labels used in the Excel file and corresponding values 	|
| user_na     	| Which values are missing values?                                            	|
| save_to     	| Directory to save the .sps file                                             	|

Example:
```R
xlsx_to_syntax(xlsx = "/Users/siardv/Documents/MIP/mip.xlsx", 
               id = 1,
               answer = 2,
               value_label = list(value = c(1:24, 991, 992, 999),
                                  label = c("Economy / Financial situation", "Social security", "Politics", "Crime", "Defense",
                                            "Healthcare", "Education", "Income / Prince levels / Taxes", "Employment",
                                            "Traffic / Mobility", "Housing", "Environment", "Population", "Minorities", 
                                            "Norms and values",  "Media", "European integration", "Inequality / Poverty", 
                                            "Intolerance / Discrimination", "Foreign policy / International security", 
                                            "Regulation / Big government", "Polarisation / Dividedness", "Immigration", 
                                            "Corona", "There are no problems", "No other problems", "DK/NA/Cannot be coded")),
               user_na = c(991, 992, 999),
               save_to = "/Users/siard/Desktop")
               
# or, to load multiple files:
xlsx = paste0("/Users/siardv/Documents/MIP/", c("mip_1", "mip_2", "mip_3"), ".xlsx")
```
