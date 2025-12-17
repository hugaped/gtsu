# Functions for creating outputs
# Author: Hugo Pedder
# Created: 2023-03-27



#' Extract coda from results and save to csv
#'
#' @param ref A treatment/class name, given as a string
#' @param writetocsv Can take `TRUE` for coda to be saved to csv, or `FALSE` for the coda
#' object to simply be refurned
#' @param scalesd A number indicating the value of a scale SD that should be used to convert SMDs back to
#' a specified scale
#' @inheritParams bugstoexcel
#'
#' @export
codatocsv <- function(outcome, ref=NULL, level="class", writetocsv=TRUE,
                      scalesd=NULL) {

  bugsmod <- readRDS(paste0("BUGSresults/", outcome, ".RDS"))

  if (level=="class") {
    codas <- bugsmod$sims.list$m
  } else if (level=="treat") {
    codas <- bugsmod$sims.list$d
  }
  codas <- cbind(rep(0,nrow(codas)), codas)

  file <- paste0("Results/", outcome, ".xlsx")

  trtcodes <- openxlsx::read.xlsx(xlsxFile=file, sheet="Intervention and Class Codes")
  trtcodes <- unique(trtcodes[[level]])
  trtcodes <- trtcodes[!is.na(trtcodes)]

  if (length(trtcodes)!=ncol(codas)) {
    stop("length(trtcodes)!=ncol(codas)")
  }
  colnames(codas) <- trtcodes

  # Rearrange and recalculate for new ref
  if (!is.null(ref)) {
    if (!ref %in% trtcodes) {
      stop("'ref' is not a named treatment/class given in the 'outcome' results Excel file")
    }

    if (trtcodes[1]!=ref) {
      index <- which(trtcodes==ref)

      codas <- codas[,c(index, 1:ncol(codas)[-index])]
      codas <- codas[,-(index+1)]

      refd <- codas[,1]

      codas <- apply(codas, MARGIN=2, FUN=function(x) (x-refd))
    }
  }

  # If returning from SMD to another scale
  if (!is.null(scalesd)) {
    codas <- codas * scalesd
  }

  if (writetocsv==TRUE) {
    write.csv(codas, file=paste0("Results/Codas/", outcome, "_", level, ".csv"), row.names = FALSE)
  }

  return(codas)
}



#' Outputs results from a BUGS model into an Excel workbook
#'
#' @param outcome A string representing an outcome name. This is used to identify the
#' correct BUGS model, and to label the outputted Excel workbook.
#' @param bugsdat An object of class `"bugsdat"` containing the info used to run the BUGS
#' model (i.e. data, treatment codes, etc.)
#' @param modnam A string representing the name of the model used
#' @param res.format Can take either `"rds"` or `"xlsx"` to indicate whether the BUGS model
#' results should be accessed as an RDS or xlsx file (xlsx uses only posterior summary
#' statistics and so is faster).
#' @param decimals The number of decimal places to which to report numerical results
#' @param addume Whether to add results for the Unrelated Mean Effects model alongside
#' direct estimates from the NMA. For these to be successfully read, a corresponding UME
#' BUGS model's results must be available and saved and an .RDS or .xlsx file (depending
#' on `res.format`) with the name `i.outcome`, where `outcome` is as specified in the
#' outcome parameter (e.g. `i.SMD_severe.xlsx`).
#' @param ref The name of the reference treatment against which the pooled treatment
#' effects will be plotted and shown in an `Effect size` worksheet.
#' @param devplot Whether to add a Dev-Dev plot showing the deviance of the UME vs NMA
#' model. To allow this, `addume` must be set to `TRUE` (see requirements for this), and
#' the `dev` node of both NMA and UME models must have been monitored.
#' @param netplot Whether to plot a network plot within a `Network Plots` worksheet
#' @param forestplot Whether to plot forest plots in the `Effect size vs reference` worksheet
#' @param rankplot Whether to plot cumulative rankograms in the `Ranks` worksheet
#' @param scalesd A number indicating the SD to use for back-transforming SMD to a
#' particular measurement scale. If left as `NULL` then no transformation will be
#' performed (i.e. leave as `NULL` if not modelling SMDs)
#' @param model.scale Indicates whether treatment effects are on `"log"` (e.g. OR, RoM) or
#' `"natural"` (e.g. mean difference) scale. If left as `NULL` (the default) this will be
#' inferred from the data.
#' @param node.names A list of names of parameters/nodes as named in the BUGS model. Each
#' element of the list should be a single character object with the name of the (monitored)
#' parameter in the BUGS model that can be read from the BUGS output. See details for info
#' on names, as well as default values.
#'
#' @details
#' For the outcome to be read in `"xlsx"` format the data in each `bugsresults.xlsx` tab (worksheet)
#' **must** have the following named variables: `node`, `mean`, `sd`, `val2.5pc`, `50%`, `val97.5pc`, `pd`, `DIC`
#'
#' Within `node.names`, element names in the list correspond to the following nodes:
#'
#' * `d="d"`: pooled relative treatment effects
#' * `m="m"`: pooled relative class effects
#' * `comps=NULL`: treatment-level mean differences / odds ratios between all treatments in the dataset.
#' The default is `NULL` in which case `bugstoexcel()` will search for nodes named either
#' `diff` or `or` and present mean differences or ORs accordingly.
#' * `compsClass=NULL`: class-level mean differences / odds ratios between all classes in the dataset.
#' Default is `NULL` in which case `bugstoexcel()` will search for nodes named either
#' ` diffClass` or `orClass` and present mean differences or ORs accordingly.
#' * `sd="sd"`: between-study SD
#' * `sdClass="sd2"`: within-class SD
#' * `rk="rk"`: treatment-level rankings
#' * `rkClass="rkClass"`: class-level rankings
#' * `dev="dev`: residual deviance contributions
#' * `totresdev="totresdev"`: total residual deviance
#'
#'
#' @export
bugstoexcel <- function(outcome="SMD_severe", bugsdat=NULL, modnam="RE random class effect",
                        res.format="rds", decimals=2, model.scale=NULL,
                        addume=TRUE, ref="Pill placebo",
                        devplot=TRUE, netplot=TRUE, forestplot=TRUE, rankplot=FALSE, scalesd=NULL,
                        node.names=NULL,
                        ...) {

  # Default node names
  def.names <- list(d="d", m="m", comps=NULL, compsClass=NULL,
                    sd="sd", sdClass="sd2", rk="rk", rkClass="rkClass",
                    dev="dev", totresdev="totresdev")

  # Replace defaults with those specified in node.names
  for (i in seq_along(node.names)) {
    def.names[[names(node.names)[i]]] <- node.names[[i]]
  }
  node.names <- def.names


  if (class(bugsdat)!="bugsdat") {
    stop("'bugsdat' must be an object of class 'bugsdat'")
  }

  writefile <- paste0("Results/", outcome, ".xlsx")
  path <- "Results/"

  wb <- openxlsx::createWorkbook(title=outcome)

  # Load NMA data
  if (res.format=="xlsx") {
    bugsres <- openxlsx::read.xlsx(xlsxFile="BUGSresults/bugsresults.xlsx", sheet=outcome)
    names(bugsres)[1] <- "node"
    bugsres <- bugsres[!is.na(bugsres$node),]

    bugsmod <- bugsres

    # Check if median column is numeric
    if (!is.numeric(bugsres$median)) {
      stop("'median' is not numeric - data header row may have been copied into the results twice")
    }

  } else if (res.format=="rds") {
    bugsmod <- readRDS(paste0("BUGSresults/", outcome, ".RDS"))

    bugsres <- as.data.frame(bugsmod$summary) %>%
      rename("median"="50%",
             "val2.5pc"="2.5%",
             "val97.5pc"="97.5%") %>%
      mutate(node=rownames(.)) %>%
      relocate(node)

  } else {
    stop("res.format specified incorrectly")
  }

  message("Loaded NMA BUGS results")

  # Check for inconsistency data
  if (addume==TRUE) {
    if (res.format=="xlsx") {

      umeres <- tryCatch({
        # The code you want run
        x <- openxlsx::read.xlsx(xlsxFile="BUGSresults/bugsresults.xlsx", sheet=paste0("i.", outcome))
      }, error = function(err) {
        NULL
      })
      if (!is.null(umeres)) {
        names(umeres)[1] <- "node"
        umeres <- umeres[!is.na(umeres$node),]

        umemod <- umeres
      } else {
        warning("UME model results not found - will not be added to output")
      }

    } else if (res.format=="rds") {

      tempfile <- paste0("BUGSresults/", paste0("i.", outcome), ".RDS")
      if (file.exists(tempfile)) {
        umemod <- readRDS(tempfile)

        umeres <- as.data.frame(umemod$summary) %>%
          rename("median"="50%",
                 "val2.5pc"="2.5%",
                 "val97.5pc"="97.5%")
        umeres$node <- rownames(umeres)

      } else {
        warning("UME model results not found - will not be added to output")
        umeres <- NULL
      }
    }

    message("Loaded UME BUGS results (if available)")
  } else {
    umeres <- NULL
  }


  # Print BUGS output
  openxlsx::addWorksheet(wb=wb, sheetName="WinBUGS Output")
  openxlsx::writeDataTable(wb=wb, sheet="WinBUGS Output", bugsres, keepNA = FALSE, rowNames = FALSE)


  # Write treatment codes, participant numbers, numbers of comparisons
  nums <- write.trtcodes(bugsdat, savefile=NULL)

  nums$codes <- nums$codes %>% group_by(class) %>% mutate(count=seq(n()))
  # nums$codes$group <- as.numeric(classcode$group[match(nums$codes$trt, classcode$trt)])
  #
  # sd2 <- bugsres$mean[grepl("^sd2", bugsres$node)]
  # nums$codes$sharevariance[nums$codes$group!=1 & nums$codes$count==1] <- as.numeric(factor(sd2, levels=unique(sd2)))
  #
  nums$codes[nums$codes$count>1,c("classcode", "class", "Nclass")] <- NA
  nums$codes$count <- NULL
  # nums$codes$group <- NULL

  openxlsx::addWorksheet(wb=wb, sheetName="Intervention and Class Codes")
  openxlsx::writeDataTable(wb=wb, sheet="Intervention and Class Codes", as.data.frame(nums$codes), rowNames = FALSE)

  openxlsx::addWorksheet(wb=wb, sheetName="N studies per comparison")
  openxlsx::writeDataTable(wb=wb, sheet="N studies per comparison", as.data.frame(nums$trt.comps), rowNames = FALSE)
  openxlsx::writeDataTable(startCol = 7, wb=wb, sheet="N studies per comparison", as.data.frame(nums$class.comps), rowNames = FALSE)

  message("Written treatment codes, participant numbers, numbers of comparisons")

  # Write BUGSdata
  openxlsx::addWorksheet(wb=wb, sheetName="Data")
  datlist <- c("arraydat", "cfb", "bf", "binom")
  row <- 1
  for (i in seq_along(datlist)) {
    if (datlist[i] %in% names(bugsdat)) {
      if (nrow(bugsdat[[datlist[i]]])>0) {
        openxlsx::writeDataTable(startRow=row, wb=wb, sheet="Data", bugsdat[[datlist[i]]], keepNA = TRUE, na.string="NA", sep=",")
        row <- row + nrow(bugsdat) + 5
      }
    }
  }

  message("Written BUGS data")


  # Write network plots tab with numbers of studies
  nstudies <- data.frame("empty"=NA)
  if ("arraydat" %in% names(bugsdat)) {
    nstudies <- cbind(nstudies, data.frame(N.Total=nrow(bugsdat$arraydat)))
  }
  if ("cfb" %in% names(bugsdat)) {
    nstudies <- cbind(nstudies, data.frame(N.CFB=nrow(bugsdat$cfb)))
  }
  if ("bf" %in% names(bugsdat)) {
    nstudies <- cbind(nstudies, data.frame(N.BF=nrow(bugsdat$bf)))
  }
  if ("bincom" %in% names(bugsdat)) {
    nstudies <- cbind(nstudies, data.frame(N.binom=nrow(bugsdat$binom)))
  }
  nstudies$empty <- NULL

  openxlsx::addWorksheet(wb=wb, sheetName="Network Plots")
  openxlsx::writeDataTable(wb=wb, sheet="Network Plots", nstudies)

  message("Written numbers of studies (network plots tab)")


  # Write model fit statistics
  modfit <- data.frame("Model"=NA, "SD"=NA, "Totresdev"=NA, "Datapoints"=NA, "pD"=NA, "DIC"=NA, "SDclass"=NA)
  modfit[1,] <- getmodfit(bugsmod=bugsmod, bugsdat=bugsdat, res.format = res.format,
                          modnam="RE – random class effect", decimals=decimals,
                          node.names=node.names)

  if (!is.null(umeres)) {
    modfit <- rbind(modfit, getmodfit(bugsmod=umemod, bugsdat=bugsdat, res.format=res.format,
                                      modnam="Inconsistency Model", decimals=decimals,
                                      node.names=node.names))
  }

  openxlsx::addWorksheet(wb=wb, sheetName="Model fit")
  openxlsx::writeDataTable(wb=wb, sheet="Model fit", modfit)

  message("Written model fit statistics")


  # TREATMENTS: Write relative effects vs DESIRED reference treatment on link scale
  if (is.null(node.names$comps)) {
    if (any(grepl("^or", bugsres$node))) {
      node.names$comps <- "or"
    } else {
      node.names$comps <- "diff"
    }
    message(paste0("Check: comps set to ", node.names$comps))
  }

  openxlsx::addWorksheet(wb=wb, sheetName="Effect Size vs Reference")
  col <- 1
  trt <- bugsdat$trtcodes$treat

  if (!ref %in% trt) {
    stop("ref is not a treatment within bugsdat$trtcodes")
  }

  if (ref!=trt[1]) {

    if (is.null(model.scale)) {
      if (node.names$comps %in% c("diff")) {
        model.scale <- "natural"
      } else if (node.names$comps=="or") {
        model.scale <- "log"
      }
    }

    # Reorder treatment effects relative to ref
    out.df <- rerefout(ref=ref, trtcodes=bugsdat$trtcodes, es=node.names$comps, bugsres=bugsres, model.scale=model.scale)
    trt <- c(ref, as.character(out.df$treat))
  } else {
    out.df <- bugsres[grepl(paste0("^", node.names$d, "\\["), bugsres$node),]
    out.df <- out.df[out.df$node!="d[1]",]
  }

  if (!is.null(scalesd)) {
    out.df[,c("val2.5pc", "median", "val97.5pc")] <- out.df[,c("val2.5pc", "median", "val97.5pc")] * scalesd
    message(paste0("Treatment effects back-transformed from SMD using a standardising SD of: ", scalesd))
  }


  if (nrow(out.df)>0) {

    if (nrow(out.df)!=length(trt[-1])) {
      stop("Number of treatment codes does not equal number of treatment effects")
    }

    out.df <- data.frame("Treatment"=trt, "Median"=c(0,round(out.df$median, decimals)),
                         "Lower95"=c(0,round(out.df$val2.5pc, decimals)),
                         "Upper95"=c(0,round(out.df$val97.5pc, decimals)),
                         "NeatOutput"=c("Reference",neatcri(out.df, decimals))
    )
    openxlsx::writeDataTable(startCol = col, wb=wb, sheet="Effect Size vs Reference", out.df)
    col <- col + ncol(out.df) + 3
    message("Written relative effects vs reference treatment")

    if (forestplot==TRUE) {
      forestplot(out.df, ylab="Treatment")
      openxlsx::insertPlot(wb, sheet="Effect Size vs Reference", startRow=2, startCol=16, height=13,
                           dpi=500)
      message("Written forest plot vs reference treatment")
    }
  } else {
    warning("Treatment vs reference effects not found")
  }



  # CLASSES Write relative effects vs DESIRED reference class on link scale
  trt <- unique(bugsdat$trtcodes$class)

  if (is.null(node.names$compsClass)) {
    if (any(grepl("^orClass", bugsres$node))) {
      node.names$compsClass <- "orClass"
      message(paste0("Check: compsClass set to ", node.names$compsClass))
    } else if (any(grepl("^diffClass", bugsres$node))) {
      node.names$compsClass <- "diffClass"
      message(paste0("Check: compsClass set to ", node.names$compsClass))
    }
  }

  if (ref!=bugsdat$trtcodes$treat[1]) {

    message("Reordering classes relative to ref")

    if (is.null(model.scale)) {
      if (node.names$compsClass=="diffClass") {
        model.scale <- "natural"
      } else if (node.names$compsClass=="orClass") {
        model.scale <- "log"
      }
    }

    # Reorder treatment effects relative to ref
    out.df <- rerefout(ref=ref, trtcodes=bugsdat$trtcodes, es=node.names$compsClass,
                       bugsres=bugsres, model.scale=model.scale)
    trt <- c(ref, as.character(out.df$treat))
  } else {
    out.df <- bugsres[grepl(paste0("^", node.names$m, "\\["), bugsres$node),]
    out.df <- out.df[out.df$node!="D[1]",]
  }

  if (!is.null(scalesd)) {
    out.df[,c("val2.5pc", "median", "val97.5pc")] <- out.df[,c("val2.5pc", "median", "val97.5pc")] * scalesd
    message(paste0("Class effects back-transformed from SMD using a standardising SD of: ", scalesd))
  }

  if (nrow(out.df)>0) {
    if (nrow(out.df)!=length(trt[-1])) {
      stop("Number of class codes does not equal number of class effects")
    }

    out.df <- data.frame("Class"=trt, "Median"=c(0, round(out.df$median, decimals)),
                         "Lower95"=c(0, round(out.df$val2.5pc, decimals)),
                         "Upper95"=c(0, round(out.df$val97.5pc, decimals)),
                         "NeatOutput"=c("Reference", neatcri(out.df, decimals))
    )
    openxlsx::writeDataTable(startCol = col, wb=wb, sheet="Effect Size vs Reference", out.df)
    message("Written relative effects vs reference class")

    if (forestplot==TRUE) {
      forestplot(out.df, ylab="Class")
      openxlsx::insertPlot(wb, sheet="Effect Size vs Reference", startRow=2, startCol=27, height=10,
                           dpi=500)
      message("Written forest plot vs reference class")
    }
  } else {
    message("Class vs reference effects (node.names$compsClass) not found")
  }



  # Write treatment rankings
  if ("decision" %in% names(bugsdat$trtcodes)) {
    trt <- bugsdat$trtcodes$treat[bugsdat$trtcodes$decision==1]
  } else {
    trt <- bugsdat$trtcodes$treat
  }

  if (!is.null(node.names$rk)) {
    openxlsx::addWorksheet(wb=wb, sheetName="Ranks")
    col <- 1
    if (any(grepl(paste0("^", node.names$rk, "\\["), bugsres$node))) {

      out.df <- bugsres[grepl(paste0("^", node.names$rk, "\\["), bugsres$node),]

      out.df <- data.frame("Treatment"=trt, "PosteriorMeanRank"=round(out.df$mean,2),
                           #"Lower95"=out.df$val2.5pc,
                           #"Upper95"=out.df$val97.5pc,
                           "Median(95%CrI)"=neatcri(out.df, decimals)
      )
      openxlsx::writeDataTable(startCol = col, wb=wb, sheet="Ranks", out.df)
      col <- ncol(out.df) + 3

      if (rankplot==TRUE) {

        g <- cumrank(outcomes=outcome, level="trt",
                     trtlist=list(as.data.frame(nums$codes))) # Add treatment codes from df

        plot(g)

        openxlsx::insertPlot(wb, sheet="Ranks", startRow=25, startCol=1,
                             dpi=500)

      }

      message("Written treatment rankings")
    } else {
      warning("No treatment rankings identified in BUGS output")
    }
  } else {
    message("node.names$rk is NULL so rankings will not be printed")
  }

  # Write class rankings
  if (!is.null(node.names$rkClass)) {
    if (any(grepl(paste0("^", node.names$rkClass, "\\["), bugsres$node))) {

      out.df <- bugsres[grepl(paste0("^", node.names$rkClass, "\\["), bugsres$node),]

      out.df <- data.frame("Class"=trt, "PosteriorMeanRank"=round(out.df$mean,2),
                           #"Lower95"=out.df$val2.5pc,
                           #"Upper95"=out.df$val97.5pc,
                           "Median(95%CrI)"=neatcri(out.df, decimals)
      )

      openxlsx::writeDataTable(startCol = col, wb=wb, sheet="Ranks", out.df)

      if (rankplot==TRUE) {

        g <- cumrank(outcomes=outcome, level="class", trtlist=list(as.data.frame(nums$codes)))

        plot(g)

        openxlsx::insertPlot(wb, sheet="Ranks", startRow=25, startCol=9,
                             dpi=500)

      }

      message("Written class rankings")
    } else {
      warning("No class rankings identified in BUGS output")
    }
  } else {
    message("node.names$rkClass is NULL so class rankings will not be printed")
  }

  # Add all TREATMENT relative effects
  if (any(grepl(paste0("^", node.names$comps, "\\["), bugsres$node))) {
    rel.df <- writeallte(bugsdat=bugsdat, bugsres=bugsres, umeres=umeres, decimals=decimals, level="treat",
                         scalesd=scalesd, node.names=node.names, model.scale=model.scale)

    openxlsx::addWorksheet(wb=wb, sheetName="Treatment Direct Effects")
    openxlsx::writeDataTable(wb=wb, sheet="Treatment Direct Effects", rel.df, keepNA=FALSE)

    if ("pvalue" %in% names(rel.df)) {
      if (node.names$comps=="or") {
        mcid <- 0.25 # CHECK THAT THIS IS the desired MCID
      } else if (node.names$comps=="diff") {
        mcid <- 0.5 # CHECK THAT THIS IS the desired MCID
      } else {
        stop("MCID cannot be ascertained or assumed from either OR or MD")
      }
      tagrow <- which(rel.df$pvalue<0.05 | (rel.df$pvalue<0.1 &
                                              abs(rel.df$Dif)>mcid))

      style <- openxlsx::createStyle(fgFill	= "#dede1d")
      openxlsx::addStyle(wb=wb, sheet="Treatment Direct Effects",
                         style=style,
                         cols=ncol(rel.df),
                         rows=tagrow + 1
      )
    }

    message("Written all direct treatment effects")
  } else {
    message("Direct treatment effects not written - node names for node.names$comps did not match BUGS output")
  }


  # Add all CLASS relative effects
  if (any(grepl(paste0("^", node.names$compsClass, "\\["), bugsres$node))) {
    rel.df <- writeallte(bugsdat=bugsdat, bugsres=bugsres, umeres=umeres, decimals=decimals, level="class",
                         scalesd=scalesd, node.names=node.names, model.scale=model.scale)

    openxlsx::addWorksheet(wb=wb, sheetName="Class Direct Effects")
    openxlsx::writeDataTable(wb=wb, sheet="Class Direct Effects", rel.df, keepNA=FALSE)

    message("Written all direct class effects")
  }

  if (devplot==TRUE & addume==TRUE) {
    g <- devplot(outcome=outcome, bugsdat=bugsdat, add.modfit=FALSE, res.format=res.format,
                 node.names=node.names, ...)

    plot(g[[1]])

    openxlsx::insertPlot(wb, sheet="Model fit", startRow=8, startCol=6 ,
                         dpi=500)

    if (nrow(g[[2]])>0) {
      openxlsx::writeDataTable(wb=wb, sheet="Model fit", startRow = 8, g[[2]], keepNA=FALSE)
    }

    message("devplot added")
  }
  if (netplot==TRUE) {

    # Treatment level network plot
    pairs.df <- bugstopairs(bugsdat, level="treat")
    netplot(pairs.df, bugsdat$trtcodes, vertex.size=5, edge.scale=0.2,
            vertex.label.dist = 2, vertex.label.cex=0.7, v.color = "class")

    openxlsx::insertPlot(wb, sheet="Network Plots", startRow=4, startCol=6,
                         dpi=500, width=12, height=8)


    # Class level network plot
    pairs.df <- bugstopairs(bugsdat, level="class")
    netplot(pairs.df, bugsdat$trtcodes, vertex.size=5, edge.scale=0.2,
            vertex.label.dist = 2, vertex.label.cex=1, shortnames=TRUE, v.color="connect")

    openxlsx::insertPlot(wb, sheet="Network Plots", startRow=40, startCol=6,
                         dpi=500, width=12, height=8)

    message("Network plots added")
  }

  openxlsx::saveWorkbook(wb=wb, file=writefile, overwrite=TRUE)

  message("FINISHED!!!!")

}






# Prepare a list of data frames of treatment/class codes with numbers of participants
# bugsdat$arraydat must contain infor on 't' and 'n'...otherwise numbers of participants cannot be
#calculated using this function
write.trtcodes <- function(bugsdat, savefile=NULL) {

  if (class(bugsdat)!="bugsdat") {
    stop("'bugsdat' must be an object of class 'bugsdat'")
  }

  trtcodes <- bugsdat$trtcodes

  if (any(grepl("^n\\[\\,[0-9]", names(bugsdat$arraydat)))) {
    nmat <- as.matrix(bugsdat$arraydat[, grepl("^n\\[\\,[0-9]", names(bugsdat$arraydat))])
    tmat <- as.matrix(bugsdat$arraydat[, grepl("^t\\[\\,[0-9]", names(bugsdat$arraydat))])
  } else {
    stop("bugsdat$arraydat must contain info on 't' and 'n'")
  }

  totn <- vector()
  for (i in seq_along(trtcodes$treatcode)) {
    totn <- append(totn, sum(nmat[tmat==trtcodes$treatcode[i]], na.rm = TRUE))
  }

  if ("class" %in% names(trtcodes)) {
    out.df <- data.frame(trtcode=trtcodes$treatcode, trt=trtcodes$treat,
                         classcode=trtcodes$classcode, class=trtcodes$class,
                         Ntrt=totn
    )
    out.df <- out.df %>% group_by(classcode) %>% mutate(Nclass=sum(Ntrt))

  } else {
    out.df <- data.frame(trtcode=trtcodes$treatcode, trt=trtcodes$treat,
                         Ntrt=totn
    )
  }

  # Calculate numbers of studies in each comparison
  comps.t <- bugstopairs(bugsdat = bugsdat, level="treat") %>%
    rename(nstudies="nr")

  if ("class" %in% names(trtcodes)) {
    comps.c <- bugstopairs(bugsdat = bugsdat, level="class") %>%
      rename(c1=t1, c2=t2, nstudies="nr")
  }

  if (!is.null(savefile)) {

    write.xlsx(as.data.frame(out.df), file=savefile, row.names = FALSE, sheetName="codes")
    write.xlsx(as.data.frame(comps.t), file=savefile, row.names = FALSE, append=TRUE, sheetName="treat.comparisons")
    write.xlsx(as.data.frame(comps.c), file=savefile, row.names = FALSE, append=TRUE, sheetName="class.comparisons")
  }

  return(list("codes"=out.df, "trt.comps"=comps.t, "class.comps"=comps.c))
}




#' Get model fit statistics from a BUGS model results
#'
#' @param bugsmod A data frame of posterior summary statistics from a BUGS model.
#'   Must include a column for pD and DIC (as well as a totresdev node).
#' @inheritParams bugstoexcel
#'
getmodfit <- function(bugsmod, bugsdat, res.format="rds", modnam="NMA", decimals=2, scalesd=NULL,
                      node.names) {
  modfit <- data.frame("Model"=modnam, "SD"=NA, "Totresdev"=NA, "Datapoints"=NA, "pD"=NA, "DIC"=NA, "SDclass"=NA)

  if (res.format=="xlsx") {
    bugsres <- bugsmod

    # if ("dic" %in% names(bugsres)) {
    #   modfit$DIC <- round(bugsres$dic[1], decimals)
    # }
    if ("pd" %in% names(bugsres)) {
      modfit$pD <- round(bugsres$pd[1], decimals)

      if (node.names$totresdev %in% bugsres$node) {
        modfit$DIC <- modfit$pD + bugsres$median[bugsres$node==node.names$totresdev]
        modfit$DIC <- round(modfit$DIC, decimals)
      }
    }

  } else if (res.format=="rds") {
    bugsres <- as.data.frame(bugsmod$summary) %>%
      rename("median"="50%",
             "val2.5pc"="2.5%",
             "val97.5pc"="97.5%") %>%
      mutate(node=rownames(.)) %>%
      relocate(node)

    if ("DIC" %in% names(bugsmod)) {
      modfit$DIC <- round(bugsmod$DIC[1], decimals)
    }
    if ("pD" %in% names(bugsmod)) {
      modfit$pD <- round(bugsmod$pD[1], decimals)

      if (node.names$totresdev %in% bugsres$node) {
        modfit$DIC <- modfit$pD + bugsres$median[bugsres$node==node.names$totresdev]
        modfit$DIC <- round(modfit$DIC, decimals)
      }
    }
  }

  if (node.names$totresdev %in% bugsres$node) {
    modfit$Totresdev <- bugsres$median[bugsres$node==node.names$totresdev]
  }
  if (!is.null(bugsdat)) {
    modfit$Datapoints <- sum(bugsdat$arraydat$`na[]`)
  }

  if (node.names$sd %in% bugsres$node) {
    message("Check: Model is a random treatment effects model")

    sdres <- bugsres[bugsres$node==node.names$sd,]
    if (!is.null(scalesd)) {
      sdres[,c("val2.5pc", "median", "val97.5pc")] <- sdres[,c("val2.5pc", "median", "val97.5pc")] * scalesd
    }

    modfit$SD <- neatcri(sdres, decimals = decimals)
  } else {
    message("Check: Model is a fixed treatment effects model (no SD)")
  }

  if (!is.null(node.names$sdClass)) {
    if (node.names$sdClass %in% bugsres$node) {
      message("Check: Random class model")

      sdres <- bugsres[bugsres$node==node.names$sdClass,]
      if (!is.null(scalesd)) {
        sdres[,c("val2.5pc", "median", "val97.5pc")] <- sdres[,c("val2.5pc", "median", "val97.5pc")] * scalesd
      }

      #sd2 <- round(sdres, decimals)
      modfit$SDclass <- neatcri(sdres, decimals = decimals)
      #modfit$SDclass <- paste(sd2, collapse=", ")
    }


  }
  return(modfit)
}



#' Write relative effects for all comparisons in the network
#'
#' @param level takes either `"treat"` or `"class"`
#' @inheritParams bugstoexcel
#'
writeallte <- function(bugsdat, bugsres, umeres=NULL, decimals=2, level="treat", scalesd=NULL,
                       node.names, model.scale="log") {

  if (level=="treat") {

    # if (any(grepl("^diff\\[", bugsres$node))) {   # Continuous results (mean differences)
    #   out.df <- bugsres[grepl("^diff\\[", bugsres$node),]
    # } else if (any(grepl("^or\\[", bugsres$node))) {    # Binary results (odds ratios)
    #   out.df <- bugsres[grepl("^or\\[", bugsres$node),]
    # }
    out.df <- bugsres[grepl(paste0("^", node.names$comps, "\\["), bugsres$node),]

    trtcodes <- bugsdat$trtcodes %>% select(treat, treatcode)
  } else if (level=="class") {

    # if (any(grepl("^diffClass\\[", bugsres$node))) {    # Continuous results (mean differences)
    #   out.df <- bugsres[grepl("^diffClass\\[", bugsres$node),]
    # } else if (any(grepl("^orClass\\[", bugsres$node))) {   # Binary results (odds ratios)
    #   out.df <- bugsres[grepl("^orClass\\[", bugsres$node),]
    # }
    out.df <- bugsres[grepl(paste0("^", node.names$compsClass, "\\["), bugsres$node),]

    trtcodes <- unique(bugsdat$trtcodes %>% select(class, classcode))
  }

  if (!is.null(scalesd)) {
    out.df[,c("val2.5pc", "median", "val97.5pc")] <- out.df[,c("val2.5pc", "median", "val97.5pc")] * scalesd
  }

  out.df$t1 <- as.numeric(gsub("([A-z]+\\[)([0-9]+)\\,([0-9]+)\\]", "\\2", out.df$node))
  out.df$t2 <- as.numeric(gsub("([A-z]+\\[)([0-9]+)\\,([0-9]+)\\]", "\\3", out.df$node))
  out.df <- arrange(out.df, t1, t2)

  out.df$trt1 <- trtcodes[[level]][out.df$t1]
  out.df$trt2 <- trtcodes[[level]][out.df$t2]

  rel.df <- data.frame("Treat1"=out.df$trt1, "Treat2"= out.df$trt2,
                       "Median"=round(out.df$median, decimals),
                       "Lower95"=round(out.df$val2.5pc, decimals),
                       "Upper95"=round(out.df$val97.5pc, decimals),
                       "NeatOutput"=neatcri(out.df, decimals)
  )
  if (level=="class") {
    names(rel.df)[1:2] <- c("Class1", "Class2")
  }

  if (!is.null(umeres)) {
    umedecimals <- decimals

    if (level=="treat") {
      ume.df <- umeres[grepl("^d\\[", umeres$node),]
    } else if (level=="class") {
      ume.df <- umeres[grepl("^m\\[", umeres$node),]
    }

    # Drop those that sample from priors
    ume.df <- ume.df[ume.df$sd<95,]

    # Invert ume results (for continuous outcomes)
    # if (any(grepl("diff", bugsres$node))) {
    if ("diff" %in% c(node.names$comps, node.names$compsClass)) {
      # Invert UME results for response and SMD
      ume.df[,c("val2.5pc", "median", "val97.5pc")] <-
        -ume.df[,c("val97.5pc", "median", "val2.5pc")]
    }

    # Convert to OR if required
    # if (grepl("^or", out.df$node[1])) {
    if (model.scale=="log") {
      ume.df[,c("val2.5pc", "median", "val97.5pc")] <- exp(ume.df[,c("val2.5pc", "median", "val97.5pc")])
      umedecimals <- decimals + 1
    } else if (model.scale=="natural") {
      umedecimals <- decimals
    } else {
      stop("model.scale must be specified for incorporating UME direct effects")
    }

    if (!is.null(scalesd)) {
      ume.df[,c("val2.5pc", "median", "val97.5pc")] <- ume.df[,c("val2.5pc", "median", "val97.5pc")] * scalesd
    }

    ume.df$t1 <- as.numeric(gsub("([A-z]+\\[)([0-9]+)\\,([0-9]+)\\]", "\\2", ume.df$node))
    ume.df$t2 <- as.numeric(gsub("([A-z]+\\[)([0-9]+)\\,([0-9]+)\\]", "\\3", ume.df$node))

    # Remove contrasts for control vs control treatments (not properly estimated in class UME model)
    # MAY NOT BE NECESSARY HERE..?
    # ctrls <- trtcodes[[paste0(level, "code")]][trtcodes$group==1]
    # ume.df <- subset(ume.df, !((t1 %in% ctrls) & (t2 %in% ctrls)))

    ume.df$MedianIncon <- round(ume.df$median, umedecimals)
    ume.df$Lower95Incon <- round(ume.df$val2.5pc, umedecimals)
    ume.df$Upper95Incon <- round(ume.df$val97.5pc, umedecimals)

    ume.df$NeatOutputIncon <- neatcri(ume.df, umedecimals, invert=FALSE)

    ume.df <- left_join(out.df, ume.df, by=c("t1", "t2"))
    rel.df <- cbind(rel.df, ume.df[,grepl("Incon", names(ume.df))])

    # Calculate variances of NMa and Direct evidence
    #if (grepl("^or", out.df$node[1])) {
    if (model.scale=="log") {
      rel.df <- rel.df %>% mutate(Var = ((log(Median) - log(Lower95)) / 1.96) ^ 2,
                                  VarIncon = ((log(MedianIncon) - log(Lower95Incon)) / 1.96) ^ 2)
    } else if (model.scale=="natural") {
      rel.df <- rel.df %>% mutate(Var = ((Median - Lower95) / 1.96) ^ 2,
                                  VarIncon = ((MedianIncon - Lower95Incon) / 1.96) ^ 2)
    }


    # Calculate Indirect estimate variance (using heuristics for if Var > VarIncon or marking as not calculable)
    # rel.df <- rel.df %>% mutate(VarInd=case_when(log(Var /VarIncon)>0.5 ~ NaN, # Not calculable
    #                                              between(log(Var / VarIncon), 0, 0.5) ~ NaN, # NMA dominated by direct estimate so inconsistency unlikely
    #                                              log(Var / VarIncon)<0 ~ 1 / ((1/Var) - (1/VarIncon))
    # ))

    rel.df <- rel.df %>% mutate(wdir=1/VarIncon,
                                #wind=1/VarInd,
                                wind=(1/Var) + wdir,
                                VarInd= 1/wind
    )

    # if (grepl("^or", out.df$node[1])) {
    if (model.scale=="log") {
      rel.df <- rel.df %>% mutate(nma=log(Median),
                                  dir=log(MedianIncon))
    } else if (model.scale=="natural") {
      rel.df <- rel.df %>% mutate(nma=(Median),
                                  dir=(MedianIncon))
    }

    rel.df <- rel.df %>% mutate(MedianInd = ((nma * (wdir + wind)) - (wdir * dir)) / wind,
                                LowerInd = MedianInd - 1.96 * (VarInd ^ 0.5),
                                UpperInd = MedianInd + 1.96 * (VarInd ^ 0.5))


    #if (grepl("^or", out.df$node[1])) {
    if (model.scale=="log") {
      rel.df <- rel.df %>%
        mutate(MedianInd = exp(MedianInd),
               LowerInd = exp(LowerInd),
               UpperInd = exp(UpperInd),
               Dif = log(MedianInd) - log(MedianIncon)
        )

    } else if (model.scale=="natural") {
      rel.df <- rel.df %>% mutate(Dif = (MedianInd) - (MedianIncon))
    }

    # Z test for ind-dir >|< 0 (assumes approximate normality)
    pval <- pnorm(rel.df$Dif / ((rel.df$VarInd + rel.df$VarIncon)^0.5))
    pval <- apply(cbind(pval, 1-pval), MARGIN=1, FUN=function(x) min(x))
    rel.df$pvalue <- pval

    rel.df <- rel.df %>% select(!Var & !VarIncon  & !wind & !wdir & !nma & !dir & !VarInd)
  }

  return(rel.df)
}





#' Returns the posterior median and upper/lower 95%CrI as a string
#'
#' @param x A data frame containing variables for the lower 95%CrI, median, and upper
#' 95%CrI of the posterior for each node
#' @param invert Whether to invert the upper and lower 95%CrI (logical)
#'
#' @export
neatcri <- function(x, decimals=2, invert=FALSE) {
  if (invert==TRUE) {
    x <- rename(x, "val2.5pc"=val97.5pc,
                "val97.5pc"=val2.5pc
    )
    #x[,2:ncol(x)] <- -x[,2:ncol(x)]
    x[,c("val2.5pc", "median", "val97.5pc")] <- -x[,c("val2.5pc", "median", "val97.5pc")]
  }

  str <- paste0(round(x$median, decimals), " (",
                round(x$val2.5pc, decimals), ", ",
                round(x$val97.5pc, decimals), ")"
  )
  return(str)
}



#' Returns the posterior median and upper/lower 95%CrI as a string **based on multinma output**
#'
#' @param x A data frame containing variables for the lower 95%CrI, median, and upper
#' 95%CrI of the posterior for each node
#' @param colnams A vector of column names corresponding to
#' c(`point estimate`, `lower interval`, `upper interval`)
#'
#' @export
neatcri2 <- function(x, colnams=c("50%", "2.5%", "97.5%"), decimals=2) {

  str <- paste0(round(x[[colnams[1]]], decimals), " (",
                round(x[[colnams[2]]], decimals), ", ",
                round(x[[colnams[3]]], decimals), ")"
  )
  return(str)
}





#' Helper function to get summary stats
get_sum <- function(x) {
  c(mean = mean(x), sd = sd(x), quantile(x, probs = c(0.025, 0.5, 0.975)))
}




#' Create posterior summaries for MCMC matrix
summary_multinma_matrix <- function(x, probs=c(0.025,0.25,0.5,0.75,0.975)) {
  p_mean <- apply(x, 2, mean)
  p_sd <- apply(x, 2, sd)
  qt <- function(x, probs, ...) {
    if (all(is.na(x)))
      setNames(rlang::rep_along(probs, NA_real_), paste0(probs *
                                                           100, "%"))
    else quantile(x, probs = probs, ...)
  }
  p_quan <- apply(x, 2, qt, probs = probs)
  p_quan <- as.data.frame(t(p_quan))

  ss <- tibble::tibble(parameter=colnames(x), mean = p_mean, sd = p_sd,
                       !!!p_quan)
  return(ss)
}




#' Writes summary BUGS data from RDS files of models to an xlsx file named "bugsresults.xlsx"
#' @param outcomes a vector of outcome names to write from rds to xlsx
#' @param overwrite indicates whether previous data should be overwritten or not
#' @param bugs.result Can be used to pass a data frame of posterior summary statistics from
#' a BUGS model directly to the function.
#'
#' @return Creates an excel file named "bugresults.xlsx" within the subdirectory BUGSresults
#'
#' @export
rdstoxlsx <- function(outcomes="dyspar", overwrite=TRUE, bugs.result=NULL) {

  file <- "BUGSresults/bugsresults.xlsx"

  wb <- openxlsx::loadWorkbook(file=file)

  for (out in seq_along(outcomes)) {

    if (is.null(bugs.result)) {
      bugsmod <- readRDS(paste0("BUGSresults/", outcomes[out], ".RDS"))
    } else {
      bugsmod <- bugs.result
    }

    bugsres <- as.data.frame(bugsmod$summary) %>%
      rename("median"="50%",
             "val2.5pc"="2.5%",
             "val97.5pc"="97.5%") %>%
      mutate(node=rownames(.)) %>%
      relocate(node)

    bugsres$dic <- c(bugsmod$DIC, rep(NA, nrow(bugsres)-1))
    bugsres$pd <- c(bugsmod$pD, rep(NA, nrow(bugsres)-1))

    if (outcomes[out] %in% openxlsx::getSheetNames(file) & overwrite==TRUE) {

      print(paste0("overwriting data for ", outcomes[out]))
      openxlsx::removeWorksheet(wb, sheet=outcomes[out])
    }

    # Create new worksheet
    openxlsx::addWorksheet(wb=wb, sheetName=outcomes[out])

    openxlsx::writeDataTable(wb=wb, sheet=outcomes[out], bugsres, keepNA = FALSE, rowNames = FALSE)

    print(paste0("written data for ", outcomes[out]))

  }

  openxlsx::saveWorkbook(wb=wb, file=file, overwrite=TRUE)
}






#' Save bugs models
#'
#' @param bugs.result An object generated from `R2OpenBUGS::bugs()`
#' @param bugsdat An object of class `"bugsdat"`
#' @param rds Whether to save a `.RDS` file of the model. If set to `TRUE` then
#' the function may take longer to run (BUGS files can be big since they contain
#' all the MCMC draws)
#' @param xlsx Whether to save a `.xlsx` file of posterior summary statistics
#'
#' @export
savemodel <- function(bugs.result, bugsdat, outcome, rds=FALSE, xlsx=TRUE) {

  # Save RDS
  if (rds==TRUE) {
    saveRDS(bugs.result, file=paste0("BUGSresults/", outcome, ".RDS"))
  }

  # Save to bugsresults.xlsx
  if (xlsx==TRUE) {
    rdstoxlsx(outcomes=outcome, overwrite = TRUE, bugs.result=bugs.result)
  }

  # Save bugsdat
  saveRDS(bugsdat, file=paste0("BUGSdata/", outcome, "_bugsdat.RDS"))
}
