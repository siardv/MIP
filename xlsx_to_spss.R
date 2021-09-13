

# Generate SPSS syntax from Excel files with pre-categorized answers to open-ended questions.

if(!"xlsx" %in% installed.packages())
  install.packages("xlsx")
library(xlsx)

`%!in%` <- Negate(`%in%`)

get_syntax <- function(df, row, id, answer, vl, na) {
  set <- df[row, which(vl[[1]][match(df[row, ], vl[[2]])] %!in%  na)][-answer]
  pid <- ifelse(any(letters %in% tolower(strsplit(set[[id]], split = "")[[1]])), sQuote(set[[id]], ""), set[id])
  if (length(set) == 1) {
    paste0("\n/*", row, ">.\nDO IF ", names(df)[id], " = ", pid, ".\nRECODE ", gsub(",", " TO", toString(range(names(df)[-c(id, answer)]))), " (SYSMIS = ", 999, ").\nEND IF.\n/*", row, "<\n\n")
  } else if (length(set) == 2) {
    paste0("\n/*", row, ">.\nIF(", names(df)[id], " = ", pid, ") ", names(set[-id]), " = ", vl[[1]][amatch(set[-id], vl[[2]])], 
           ". /* ", dQuote(tolower(gsub("/", "-", set[-id])), q = "double"), ".\n/*", row, "<\n\n")
  } else if (length(set) > 2) {
    paste0("\n/*", row, ">.\nDO REPEAT\nX = ", toString(names(set[-id])),"\n/Y = ", toString(vl[[1]][amatch(set[-id], vl[[2]])]),
           ". /* ", toString(dQuote(tolower(gsub("/", "-", set[-id])), q = "double")),
           ".\nDO IF ", names(set)[id], " = ", pid, ".\nRECODE X (SYSMIS = Y).\nEND IF.\nEND REPEAT.\n/*", row, "<\n\n")
  }
}

xlsx_to_syntax <- function(xlsx, id, answer, value_label, user_na, save_to){
  sapply(xlsx, function(x) {
    message("Creating syntax...")
    df <- read.xlsx(x, sheetIndex = 1)
    df <- as.data.frame(apply(df, 2, function(x) gsub("^$", "No other problems", gsub("IC|DK|Cannot.*$|NA", "DK/NA/Cannot be coded", ifelse(is.na(x), "NA", x)))))
    names(df) <- c(names(df)[id], names(df)[answer], paste0(names(df)[2], letters[1:(ncol(df) - 2)]))
    sps <- paste0(gsub("^.*/|[.].*$", "", x), ".sps")
    sink(sps, append = FALSE)
    cat(paste0("/* syntax for ", gsub("^.*/", "~/", x), " (n = ", nrow(df), ").\n\n"))
    cat(paste0("GET  FILE = '", x,"'.\nDATASET NAME DEMES WINDOW = FRONT.\n\n"))
    probs <- names(df)[-c(id, answer)]
    cat(paste0("DO REPEAT\nX = ", toString(probs), ".\nCOMPUTE  X = $SYSMIS.\nFORMATS X (F3).\nMISSING VALUES X (", toString(user_na), ").\nEND REPEAT.\n\n"))
    cat("VALUE LABELS ", probs[1], " TO " , tail(probs, 1), "\n")
    cat(paste0(gsub(", ", "\n", toString(paste0(vl[[1]], " '", vl[[2]], "'"))), ".\nEXECUTE.\n"))
    cat(unlist(sapply(1:nrow(df), function(i) unlist(ifelse(i == nrow(df), 
                                                            list(paste0(get_syntax(df, i, id, answer, value_label, user_na), 
                                                                        paste0("\nDO REPEAT ID = ", names(df)[id], ".\nDO IF NOT (SYSMIS(", probs[1],")) AND (", 
                                                                               names(df)[id]," = ID).\nRECODE ", probs[1], " TO ", tail(probs[1], 1), " (SYSMIS = ", 992, ").\nELSE IF (SYSMIS(", probs[1],")) AND (", 
                                                                               names(df)[id]," = ID).\nRECODE ", probs[1], " TO ", tail(probs[1], 1), " (SYSMIS = ", 999, ").\nEND IF.\nEND REPEAT.\nEXECUTE.\n"))), 
                                                            list(get_syntax(df, i, id, answer, value_label, user_na)))))))
    sink()
    file.copy(sps, to = save_to)
    message(paste0("Done! Saved to ", path.expand("~"), save_to))
  })
}




## Arguments     
# xlsx            Directory to Excel file(s) containing pre-categorized answers
# id              Individual-level identifier
# answer          Original answer to open-ended question
# value_label     List containing categories/labels used in the Excel file with corresponding values
# user_na         Which values are missing values?
# save_to         Directory to save the .sps file
