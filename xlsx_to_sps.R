xlsx_to_sps <- function(path, id, answer, value_labels, user_na, save_to){
    `%!in%` <- Negate(`%in%`)
    get_syntax <- function(df, row, id, answer, vl, na) {
        set <- df[row, which(vl[[1]][match(df[row, ], vl[[2]])] %!in%  na)][-answer]
        pid <- ifelse(any(letters %in% tolower(strsplit(set[[id]], split = "")[[1]])), sQuote(set[[id]], ""), set[id])
        if (length(set) == 1) {
            paste0("\n/*", row, ">.\nDO IF ", names(df)[id], " = ", pid, ".\nRECODE ", gsub(",", " TO", toString(range(names(df)[-c(id, answer)]))), " (SYSMIS = ", 999, ").\nEND IF.\n/*", row, "<\n")
        } else if (length(set) == 2) {
            paste0("\n/*", row, ">.\nIF(", names(df)[id], " = ", pid, ") ", names(set[-id]), " = ", vl[[1]][match(set[-id], vl[[2]])], 
                   ". /* ", dQuote(tolower(gsub("/", "-", set[-id])), q = "double"), ".\n/*", row, "<\n")
        } else if (length(set) > 2) {
            paste0("\n/*", row, ">.\nDO REPEAT\nX = ", toString(names(set[-id])),"\n/Y = ", toString(vl[[1]][match(set[-id], vl[[2]])]),
                   ". /* ", toString(dQuote(tolower(gsub("/", "-", set[-id])), q = "double")),
                   ".\nDO IF ", names(set)[id], " = ", pid, ".\nRECODE X (SYSMIS = Y).\nEND IF.\nEND REPEAT.\n/*", row, "<\n")
        }
    }
    invisible(sapply(seq_along(path), function(i) {
        message("Creating syntax...")
        df <- read.xlsx(path[i], sheetIndex = 1)
        df <- as.data.frame(apply(df, 2, function(i) gsub("^$", "No other problems", gsub("IC|DK|Cannot.*$|NA", "DK/NA/Cannot be coded", ifelse(is.na(i), "NA",i)))))
        names(df) <- c(names(df)[id], names(df)[answer], paste0(names(df)[answer], letters[seq(1, ncol(df) - 2)]))
        txt <- paste0(gsub("^.*/|[.].*$", "", path[i]), ".txt")
        probs <- names(df)[-c(id, answer)]
        cat(paste0("/* syntax for ", gsub("^.*/", "~/", path[i]), " (n = ", nrow(df), ").\n"),
            paste0("\rGET  FILE = '[PATH TO .SAV FILE]'.\nDATASET NAME DEMES WINDOW = FRONT.\n"),
            paste0("\rDO REPEAT\nX = ", toString(probs), ".\nCOMPUTE X = $SYSMIS.\nFORMATS X (F3).\nMISSING VALUES X (", toString(user_na), ").\nEND REPEAT.\n"),
            paste0("\rVALUE LABELS ", probs[1], " TO " , tail(probs, 1), "\n"),
            paste0(gsub(", ", "\n", toString(paste0(value_labels[[1]], " '", value_labels[[2]], "'"))), ".\nEXECUTE.\n"),
            sapply(1:nrow(df), function(r) unlist(ifelse(r == nrow(df), 
                                                         list(paste0(get_syntax(df, r, id, answer, value_labels, user_na), 
                                                                     paste0("\nDO REPEAT ID = ", names(df)[id], ".\nDO IF NOT (SYSMIS(", probs[1],")) AND (", 
                                                                            names(df)[id]," = ID).\nRECODE ", probs[1], " TO ", tail(probs, 1), " (SYSMIS = ", 992, ").\nELSE IF (SYSMIS(", probs[1],")) AND (", 
                                                                            names(df)[id]," = ID).\nRECODE ", probs[1], " TO ", tail(probs, 1), " (SYSMIS = ", 999, ").\nEND IF.\nEND REPEAT.\n\nEXECUTE.\n"))), 
                                                         list(get_syntax(df, r, id, answer, value_labels, user_na))))), file = txt, append = TRUE)
        sapply(gsub(".txt", ".sps", txt), function(sps) {
            file.rename(txt, sps)
            file.copy(sps, save_to)
            message(paste0("Done! Saved ", sps, " to ", path.expand("~"), save_to))
        })
    }))
}
