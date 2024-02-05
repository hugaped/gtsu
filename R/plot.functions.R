# Functions for plotting
# Author: Hugo Pedder
# Created: 2023-03-27


#' Plots a basic forest plot for data frame of outcomes
#'
#' Use within internal `gtsu` functions, rather than by the user
#'
forestplot <- function(out.df, refline=TRUE, refvalue=0, xlab="Effect size",
                       ylab="Treatment/Class") {

  names(out.df)[1] <- "labs"

  g <- ggplot(out.df, aes(x=labs, y=Median, ymin=Lower95, ymax=Upper95), height=2) +
    geom_point() +
    geom_errorbar(linewidth=0.5, width=0.5) +
    coord_flip() +
    xlab(ylab) + ylab(xlab)

  if (refline==TRUE) {
    g <- g + geom_hline(yintercept=refvalue, linetype="dashed", color="red")
  }

  g <- g + theme_bw()

  plot(g)
}



#' Plots a dev-dev plot
#'
#' @param outcome A character object specifying the name of the analysis for which to plot a dev-dev plot.
#' Must match a named worksheet in `"BUGSresults/bugsresults.xlsx"`, and there must be a corresponding UME model
#' that has been run and saved with the `i.` prefix. Note that `dev` nodes
#' **must have been monitored and saved in both NMA and UME models to be able to plot a dev-dev plot**.
#' @param bugsdat An object of class `"bugsdat"` that includes the data used in the NMA. This will eventually
#' be deprecated to ensure it is stored elsewhere and doesn't need to be formally specified within this function.
#' @param res.format Can be `"xlsx"` or `"rds"` to indicate the format in which BUGS results have been saved.
#' The `savemodel()` function now saves results in both formats, and loading from xlsx is much more rapid so
#' leaving as the default is recommended.
#' @param lim A numeric vector of length 2 specifying the limits for the axes
#' @param add.modfit Indicates whether or not to include model fit statistics as annotations on the plot. Messy
#' but informative. Though the relevant info must have been monitored and saved.
#' @param ratcut A number representing the cutoff for the ratio of nmadev/umedev, above which (and if they also
#' exceed the value specified in `devcut`) points should be coloured red and their values returned in a
#' data frame.
#' @param devcut A number representing the cutoff value for the deviance, above which (and if they also
#' exceed the value specified in `ratcut`) points should be coloured red and their values returned in a data frame.
#' @param direction If given as `"nma"` it specifies `ratcut` should be calculated as nmadev/umedev and `devcut`
#' should be a threshold for the nma deviance. If given as `"ume"` it specifies `ratcut` should be calculated as
#' umedev/nmadev and `devcut` should be a threshold for the ume deviance.
#' @inheritParams bugstoexcel
#'
#' @export
devplot <- function(outcome, bugsdat=NULL, res.format="xlsx",
                    lim=NULL, add.modfit=TRUE,
                    ratcut=0, devcut=5, direction="nma",
                    node.names=list(d="d", m="m", comps=NULL, compsClass=NULL,
                                    sd="sd", sdClass="sd2", rk="rk", rkClass="rkClass",
                                    dev="dev", totresdev="totresdev")
                    #priority="cfb", # from depression
                    #save=NULL
                    ) {

  if (res.format=="xlsx") {
    bugsres <- readxl::read_xlsx("BUGSresults/bugsresults.xlsx", sheet=outcome)
    names(bugsres)[1] <- "node"
    bugsmod <- bugsres

    umeres <- readxl::read_xlsx("BUGSresults/bugsresults.xlsx", sheet=paste0("i.", outcome))
    names(umeres)[1] <- "node"
    umemod <- umeres

  } else if (res.format=="rds") {

    # NMA
    bugsmod <- readRDS(paste0("BUGSresults/", outcome, ".RDS"))

    bugsres <- as.data.frame(bugsmod$summary) %>%
      rename("median"="50%",
             "val2.5pc"="2.5%",
             "val97.5pc"="97.5%") %>%
      mutate(node=rownames(.)) %>%
      relocate(node)

    # UME
    umemod <- readRDS(paste0("BUGSresults/i.", outcome, ".RDS"))

    umeres <- as.data.frame(umemod$summary) %>%
      rename("median"="50%",
             "val2.5pc"="2.5%",
             "val97.5pc"="97.5%")
    umeres$node <- rownames(umeres)

  }

  matchdev <- paste0("^", node.names$dev, "\\[")

  if (!any(grepl(matchdev, bugsres$node))) {
    stop("'dev' not monitored in NMA model")
  }
  if (!any(grepl(matchdev, umeres$node))) {
    stop("'dev' not monitored in UME model")
  }

  devs <- data.frame("node"=bugsres$node[grepl(matchdev, bugsres$node)],
                     "nma"=bugsres$mean[grepl(matchdev, bugsres$node)],
                     "ume"=umeres$mean[grepl(matchdev, umeres$node)]
  )

  if (direction=="nma") {
    devs$ratio <- devs$nma / devs$ume
    devs$outlier <- ifelse(devs$ratio>=ratcut & devs$nma>=devcut,1,0)
  } else if (direction=="ume") {
    devs$ratio <- devs$ume / devs$nma
    devs$outlier <- ifelse(devs$ratio>=ratcut & devs$ume>=devcut,1,0)
  }


  if (is.null(lim)) {
    lim <- c(0,max(c(devs$nma, devs$ume)))
  }

  g <- ggplot(data=devs, aes(x=nma, y=ume, color=factor(outlier))) +
    geom_point() +
    scale_color_manual(values=c("black", "red")) +
    xlim(lim) + ylim(lim) +
    xlab("NMA residual deviances") + ylab("Inconsistency residual deviances")

  if (add.modfit==TRUE) {

    nma.modfit <- getmodfit(bugsmod=bugsmod, bugsdat=bugsdat, res.format = res.format,
                            modnam="RE – random class effect", decimals=2,
                            node.names=node.names)

    ume.modfit <- getmodfit(bugsmod=umemod, bugsdat=bugsdat, res.format = res.format,
                            modnam="UME", decimals=2,
                            node.names=node.names)

    g <- g +
      geom_text(x=lim[2]/6, y=(lim[2]*3)/4, label=paste0("UME totresdev: ", ume.modfit$Totresdev, "\n",
                                                         "UME DIC: ", ume.modfit$DIC, "\n",
                                                         "UME between-study SD: ", ume.modfit$SD,"\n",
                                                         "N Data Points: ", nrow(devs)),
                hjust = 0
      ) +
      geom_text(y=lim[2]/6, x=(lim[2]*3)/4, label=paste0("NMA totresdev: ", nma.modfit$Totresdev, "\n",
                                                         "NMA DIC: ", nma.modfit$DIC, "\n",
                                                         "NMA between-study SD: ", nma.modfit$SD,"\n",
                                                         "N Data Points: ", nrow(devs)),
                hjust = 0
      )

  }
  if (all(devs$outlier==0) | add.modfit==FALSE) {
    g <- g +
      theme(legend.position = "none")
  }

  if (!is.null(bugsdat)) {

    if (class(bugsdat)!="bugsdat") {
      stop("'bugsdat' must be an object of class 'bugsdat'")
    }

    studies <- vector()
    studies <- append(studies, bugsdat$arraydat$studyid)

    # From Depression
    # if (priority=="binom") {
    #   studies <- append(studies, bugsdat$binom$studyid)
    # }
    # studies <- append(studies, bugsdat$cfb$studyid)
    # studies <- append(studies, bugsdat$bf$studyid)
    # if (priority=="cfb") {
    #   studies <- append(studies, bugsdat$binom$studyid)
    # }


    out.df <- data.frame("devs"=bugsres$node[match(devs$node, bugsres$node)][which(devs$outlier==1)])
    out.df$studies <- studies[as.numeric(gsub("(dev\\[)([0-9]+)(\\,[0-9]\\])", "\\2", out.df$devs))]
    out.df$nmadev <- devs$nma[which(devs$outlier==1)]

    if (nrow(out.df)>0) {
      print(out.df)
    } else {
      print("No data points have deviances exceeding 'devcut' or a deviance ratio (nma/ume) exceeding 'ratcut'")
    }
  } else {
    print(bugsres$node[which(devs$outlier==1)])
  }

  # if(!is.null(save)) {
  #   g <- g + ggtitle(nma)
  #
  #   officer::read_pptx(save) %>%
  #     officer::add_slide() %>%
  #     officer::ph_with(g, ph_location_fullsize()) %>%
  #     base::print(target = save)
  # }

  return(list(g, out.df))
}






# Plots forest plots using info from compiled results files
# `ref` must be a named treatment in all datasets
# Outcome can be "diff", "diffClass", "or" or "orClass"

#' Plots results from multiple NMAs on same forest plot
#'
#' Useful for showing sensitivity/subgroup anlaysis results
#'
#' @param analyses A character vector of analyses to plot on the same forest plot. Elements must be the same
#' as names for workbooks in `"Results"` directory.
#' @param ref A named treatment/class to use as the reference (against which relative effects will be plotted).
#' Must be present in all analysis results given in `analyses` and must match the treatment/class level
#' indicated by `param`.
#' @param refline Whether to plot a vertical dashed line indicating the reference treatment effect
#' @param param A character object indicating the relative effect parameter from the NMA to be plotted.
#' Within this is indicated whether effects are plotted at class or treatment level. Typical options are `diff`,
#' `diffClass`, `or`, `orClass`.
#' @param analysis.lab A character vector of the same length as `analyses` indicating the labels for each
#' analysis to be shown in the legend. Can use `\n` for a new line if the label should be on multiple lines.
#' @param dodge A number indicating how far apart to plot results from differnt analyses
#' @param caption A character object specifying a caption to include underneath the plot.
#'
#' @export
forestsense <- function(analyses=c("SMD_severe"), ref="Pill placebo", refline=TRUE, param="diff",
                        analysis.lab=NULL,
                        dodge=0.5, caption="") {

  file <- paste0("Results/", analyses[1], ".xlsx")
  plot.df <- openxlsx::read.xlsx(xlsxFile=file, sheet="WinBUGS Output") %>%
    mutate(treat=NA, analysis=NA)

  # If BUGS output has dic and pd
  plot.df$pd <- NULL
  plot.df$dic <- NULL

  plot.df <- plot.df[0,]
  for (i in seq_along(analyses)) {
    file <- paste0("Results/", analyses[i], ".xlsx")
    bugsres <- openxlsx::read.xlsx(xlsxFile=file, sheet="WinBUGS Output")
    trtcodes <- openxlsx::read.xlsx(xlsxFile=file, sheet="Intervention and Class Codes")

    # If BUGS output has dic and pd
    bugsres$pd <- NULL
    bugsres$dic <- NULL

    out.res <- subset(bugsres, grepl(paste0(param, "\\["), node))

    # Change reference treatment by switching around ORs
    if (grepl("Class", param) | param %in% "m") {
      lvl <- "class"
      level <- "class"

    } else {
      lvl <- "trt"
      level <- "treat"
    }

    if (trtcodes[[lvl]][trtcodes[[paste0(lvl, "code")]] %in% 1]!=ref) {

      if (!ref %in% trtcodes[[lvl]]) {
        stop("'ref' not a common treatment/class in all datasets")
      }

      index <- trtcodes[[paste0(lvl, "code")]][which(trtcodes[[lvl]]==ref)]
      if (length(index)>1) {
        stop("Index class must be a class with only a single treatment, where treatment=class")
      }

      upper <- subset(out.res, grepl(paste0(param, "\\[", index, "\\,"), node)) %>%
        # Ensure no indices lower than index are included in upper
        mutate(temp=gsub(paste0(param, "\\[", index, "\\,([0-9]+)\\]"), "\\1", node)) %>%
        mutate(temp=as.numeric(temp)) %>%
        subset(temp>index) %>% select(!temp)

      upper$treat <- trtcodes[[lvl]][!is.na(trtcodes[[paste0(lvl, "code")]]) &
                                       trtcodes[[paste0(lvl, "code")]]>index]

      lower <- subset(out.res, grepl(paste0(param, "\\[[0-9]+\\,", index, "\\]"), node))

      # Inverst result
      if (grepl("or", param)) {
        lower <- lower %>% mutate(tmedian=exp(-log(median)),
                                  tval2.5pc=exp(-log(val97.5pc)),
                                  tval97.5pc=exp(-log(val2.5pc))
        ) %>%
          mutate(median=tmedian,
                 val2.5pc=tval2.5pc,
                 val97.5pc=tval97.5pc
          ) %>% select(!c(tmedian, tval2.5pc, tval97.5pc))
      } else {
        lower <- lower %>% mutate(tmedian=-median,
                                  tval2.5pc=-val97.5pc,
                                  tval97.5pc=-val2.5pc
        ) %>%
          mutate(median=tmedian,
                 val2.5pc=tval2.5pc,
                 val97.5pc=tval97.5pc
          ) %>% select(!c(tmedian, tval2.5pc, tval97.5pc))
      }
      lower <- lower %>% mutate(treat = gsub(paste0(param, "\\[([0-9]+)\\,[0-9]+\\]"), "\\1", lower$node)) %>%
        mutate(treat=trtcodes[[lvl]][which(trtcodes[[paste0(lvl, "code")]] %in% treat)])


      codes <- unique(trtcodes[[lvl]])
      codes <- codes[!is.na(codes)]
      codes <- codes[codes!=ref]

      out.res <- rbind(upper, lower) %>%
        mutate(treat=factor(treat, levels=codes))

      # Convert OR to logOR
      if (grepl("or", param)) {
        out.res <- out.res %>% mutate(tmedian=log(median),
                                      tval2.5pc=log(val97.5pc),
                                      tval97.5pc=log(val2.5pc)
        ) %>%
          mutate(median=tmedian,
                 val2.5pc=tval2.5pc,
                 val97.5pc=tval97.5pc
          ) %>% select(!c(tmedian, tval2.5pc, tval97.5pc))
      }

      if (is.null(analysis.lab)) {
        if (grepl("pharm", analyses[i])) {
          out.res$analysis <- "Non-Pharma Only"
        } else if (grepl("bias", analyses[i])) {
          out.res$analysis <- "Bias-Adjusted"
        } else if (grepl("sensipt", analyses[i])) {
          out.res$analysis <- "Class change"
        } else if (grepl("time", analyses[i])) {
          out.res$analysis <- "Study duration adjusted"
        } else if (grepl("sensrob", analyses[i])) {
          out.res$analysis <- "Low risk Attrition"
        } else if (grepl("sensitt", analyses[i])) {
          out.res$analysis <- "ITT studies only"
        } else if (grepl("sensshort", analyses[i])) {
          out.res$analysis <- "Long term (>12 week) studies excluded"
        } else {
          out.res$analysis <- "Base-Case NMA"
        }
      } else {
        out.res$analysis <- analysis.lab[i]
      }

      plot.df <- rbind(plot.df, out.res)

    } else {

      if (param %in% c("d", "m")) {

        out.res <- subset(out.res, grepl(paste0("^", param, "\\["), node))

        out.res <- out.res %>% mutate(treat = gsub(paste0("^", param, "\\[([0-9]+)\\]"), "\\1", out.res$node)) %>%
          mutate(treat = as.numeric(treat)) %>%
          arrange(treat)

      } else {
        out.res <- subset(out.res, grepl(paste0(param, "\\[1\\,"), node))

        out.res <- out.res %>% mutate(treat = gsub(paste0(param, "\\[[0-9]+\\,([0-9]+)\\]"), "\\1", out.res$node)) %>%
          mutate(treat = as.numeric(treat)) %>%
          arrange(treat)
      }

      codes <- unique(trtcodes[[lvl]])[-1]
      codes <- codes[!is.na(codes)]
      out.res <- out.res %>%
        mutate(treat=factor(codes, levels=codes))

      # Convert OR to logOR
      if (grepl("or", param)) {
        out.res <- out.res %>% mutate(median=log(median),
                                      val2.5pc=log(val2.5pc),
                                      val97.5pc=log(val97.5pc)
        )
      }

      if (is.null(analysis.lab)) {
        if (grepl("pharm", analyses[i])) {
          out.res$analysis <- "Non-Pharma Only"
        } else if (grepl("bias", analyses[i])) {
          out.res$analysis <- "Bias-Adjusted"
        } else if (grepl("sensipt", analyses[i])) {
          out.res$analysis <- "Class change"
        } else if (grepl("time", analyses[i])) {
          out.res$analysis <- "Study duration adjusted"
        } else if (grepl("sensrob", analyses[i])) {
          out.res$analysis <- "Low risk Attrition"
        } else if (grepl("sensitt", analyses[i])) {
          out.res$analysis <- "ITT studies only"
        } else if (grepl("sensshort", analyses[i])) {
          out.res$analysis <- "Long term (>12 week) studies excluded"
        } else {
          out.res$analysis <- "Base-Case NMA"
        }
      } else {
        out.res$analysis <- analysis.lab[i]
      }

      plot.df <- rbind(plot.df, out.res)
    }

  }

  # Factor analysis
  if (is.null(analysis.lab)) {
    plot.df$Analysis <- factor(plot.df$analysis)
  } else {
    plot.df$Analysis <- factor(plot.df$analysis, levels=analysis.lab)
  }

  cols <- RColorBrewer::brewer.pal(3, "Set1")[c(2,1,3)]
  dodge <- position_dodge(width=dodge)

  g <- ggplot(plot.df, aes(y=treat, x=median, xmin=val2.5pc, xmax=val97.5pc,
                           color=Analysis, linetype=Analysis), height=2) +
    geom_point(position = dodge) +
    geom_errorbar(size=0.5, width=0.5, position = dodge) +
    #coord_flip() +
    ylab("Treatment/Class") + xlab("Effect size") +
    scale_color_manual(values=cols)

  if (refline==TRUE) {
    g <- g + geom_vline(xintercept=0, linetype="dashed", color="black")
  }

  g <- g + labs(caption=caption)

  g <- g + theme_bw()

  return(g)
}






#' Plots cumulative rankograms
#'
#' @param outcomes A character vector of outcome names corresponding to worksheet names
#' in `"BUGSresults/bugsresults.xlsx"` *or* Results workbooks in `"Results"` directory.
#' If multiple outcomes are specified then function plots a
#' separate panel for each treatment/class with multiple outcomes. Otherwise plots a single
#' outcome panel with all treatments/classes on the same plot.
#' @param level Can be `"trt"` for treatment-level results, or `"class"` for class-leve reuslts
#' @param byoutcome Indicates whether panels should be plotted by outcome (`TRUE`) or treatment/class
#' @param trtlist Used primarily within the `bugstoexcel()` to allow naming of treatments/classes.
#' A list in which each element is a set of treatment/class names for each outcome in `outcomes`. If left
#' as `NULL` (the default) then the function will load ranking results from outcome workbooks in `Results`
#' directory.
#' @param out.labs A character vector of labels for each of `outcomes`. Must be same length as `outcomes`
#'
#' @export
cumrank <- function(outcomes, level="trt", byoutcome=TRUE, trtlist=list(),
                    out.labs=outcomes) {

  # file <- paste0("Results/", outcomes[1], ".xlsx")
  # plot.df <- openxlsx::read.xlsx(xlsxFile=file, sheet="WinBUGS Output") %>%
  #   mutate(treat=NA, outcome=NA, rk=NA, cumsum=NA)

  if (length(out.labs)!=length(outcomes)) {
    stop("length(out.labs) must equal length(outcomes)")
  }

  file <- paste0("BUGSresults/bugsresults.xlsx")
  plot.df <- openxlsx::read.xlsx(xlsxFile=file, sheet=outcomes[1]) %>%
    mutate(treat=NA, outcome=NA, rk=NA, cumsum=NA)

  # If BUGS output has dic and pd
  plot.df$pd <- NULL
  plot.df$dic <- NULL

  plot.df <- plot.df[0,]
  for (i in seq_along(outcomes)) {

    file <- paste0("BUGSresults/bugsresults.xlsx")
    bugsres <- openxlsx::read.xlsx(xlsxFile=file, sheet=outcomes[1])

    if (length(trtlist)==0) {
      file <- paste0("Results/", outcomes[i], ".xlsx")
      trtcodes <- openxlsx::read.xlsx(xlsxFile=file, sheet="Intervention and Class Codes")
      trtcodes <- unique(trtcodes[[level]])
      trtcodes <- trtcodes[!is.na(trtcodes)]
    } else {
      trtcodes <- trtlist[[i]]
      trtcodes <- unique(trtcodes[[level]])
      trtcodes <- trtcodes[!is.na(trtcodes)]
    }


    # If BUGS output has dic and pd
    bugsres$pd <- NULL
    bugsres$dic <- NULL

    if (level=="class") {
      out.res <- subset(bugsres, grepl("probClass\\[", node))
    } else if (level=="trt") {
      out.res <- subset(bugsres, grepl("prob\\[", node))
    }

    index.trt <- as.numeric(gsub(".+\\[[0-9]+\\,([0-9]+)\\]", "\\1", out.res$node))
    out.res$rk <- as.numeric(gsub(".+\\[([0-9]+)\\,[0-9]+\\]", "\\1", out.res$node))

    out.res$treat <- trtcodes[index.trt]

    # Calculate cumulative probability ranking
    out.res <- out.res %>% group_by(treat) %>%
      mutate(cumprob=cumsum(mean))

    out.res$outcome <- outcomes[i]

    plot.df <- rbind(plot.df, out.res)
  }

  plot.df$treat <- factor(plot.df$treat)
  plot.df$outcome <- factor(plot.df$outcome, levels=outcomes,
                            labels=out.labs)

  if (n_distinct(plot.df$treat)<=9) {
    cols <- RColorBrewer::brewer.pal(n_distinct(plot.df$treat), "Set1")
  } else {
    cols <- RColorBrewer::brewer.pal(9, "Set1")

    cols <- append(cols, RColorBrewer::brewer.pal(n_distinct(plot.df$treat)-9, "Set2"))
  }



  if (byoutcome==TRUE) {
    g <- ggplot(plot.df, aes(x=rk, y=cumprob, linetype=treat, colour=treat)) +
      geom_line(size=0.8) +
      scale_color_manual(values=cols, name=NULL) +
      scale_linetype(name=NULL)+
      xlab("Rank (1=best)") + ylab("Cumulative probability") +
      theme_bw() +

      facet_wrap(~outcome)

  } else {
    g <- ggplot(plot.df, aes(x=rk, y=cumprob, linetype=outcome, colour=outcome)) +
      geom_line(size=0.8) +
      scale_color_manual(values=cols, name=NULL) +
      scale_linetype(name=NULL)+
      xlab("Rank (1=best)") + ylab("Cumulative probability") +
      theme_bw() +

      facet_wrap(~treat)
  }

  return(g)
}





#' Plots network plot to a powerpoint file (editable)
#'
#' Note that this actually isn't necessary to produce an editable plot since a plot can instead just be saved as
#' an .emf file (metafile format) (e.g. using `devEMF::emf()`), which when opened in powerpoint is editable.
#'
#' @export
netpptx <- function(bugsdat, vertex.label.cex=0.2, t.v.scale=NULL, c.v.scale=NULL,
                    file="Plots/network_discany.pptx",
                    title="Discontinuation due to any reason"
) {

  # Moderate-severe network plot - treatment-level
  main <- paste0("Treatment-level network: ", title)

  # Network plot
  pairs.df <- bugstopairs(bugsdat = bugsdat,
                          level="treat")

  g <- rvg::dml(code= netplot(pairs.df, bugsdat$trtcodes, v.color = "class",
                              edge.scale=2, vertex.label.dist = 1.7, v.scale=t.v.scale, vertex.label.cex=1,
                              main=main))

  officer::read_pptx() %>%
    officer::add_slide() %>%
    officer::ph_with(g, officer::ph_location_fullsize()) %>%
    base::print(target = file)


  # Moderate-severe network plot - class-level
  main <- paste0("Class-level network: ", title)

  # Network plot
  pairs.df <- bugstopairs(bugsdat = bugsdat,
                          level="class")

  g <- rvg::dml(code = netplot(pairs.df, bugsdat$trtcodes, v.color = "connect",
                               edge.scale=2, vertex.label.dist = 1.7, v.scale=c.v.scale, vertex.label.cex=1,
                               main=main))

  officer::read_pptx(file) %>%
    officer::add_slide() %>%
    officer::ph_with(g, officer::ph_location_fullsize()) %>%
    base::print(target = file)

}






#' Plot network plot
#'
#' Takes a dataset of contrasts and generates a network plot
#'
#' @param plotdat A dataframe of contrasts containing `t1` and `t2` for treatment/class names/codes,
#' and `nr` representing the number of studies that include this contrast. Can easily be generated using
#' `bugstopairs()`.
#' @param classcode An object of class `"classcode"`, which is a dataframe of treatment and class names and codes
#' @param v.color Can take `"connect"` to colour only nodes that are connected to the network or `"class"` to colour by class
#' @param level Can take `"trt"` to plot nodes at the treatment level or `"class"` to plot nodes at the class level
#' @param remove.loops A logical object to indicate whether loops that compare within a node should be plotted
#' @param legend Indicates whether a legend of class names should be plotted or not...sometimes setting to `TRUE` can
#' lead to difficult sizing of the network plot
#' @param edge.scale A number by which to scale the connecting (edge) widths. Can leave as `NULL` for all edges to
#' be plotted with the same width
#' @param v.scale A number by which to scale the vertex (node) size. Can leave as `NULL` for all vertices to
#' be plotted with the same size
#' @param vertex.label.dist A number indicating how far vertex (node) labels should be plotted from the vertices (nodes)
#' @param shortnames Can be set to true to label vertices (nodes) using `shorttreatment` or `shortclass` variables in `classcodes`
#' @param ... Arguments to be sent to `igraph`
#'
#' @export
netplot <- function(plotdat, classcode, v.color="connect", level="trt", remove.loops=TRUE, legend=TRUE,
                    edge.scale=NULL, v.scale=NULL, vertex.label.dist=0, shortnames=FALSE,
                    ...) {

  # Run checks
  argcheck <- checkmate::makeAssertCollection()
  # Add class checks for plotdat and classcode
  checkmate::assertChoice(v.color, choices=c("connect", "class"), add=argcheck)
  checkmate::assertChoice(level, choices=c("trt", "class"), add=argcheck)
  checkmate::assertLogical(remove.loops, add=argcheck)
  checkmate::assertLogical(legend, add=argcheck)
  checkmate::assertLogical(shortnames, add=argcheck)
  checkmate::assertNumber(edge.scale, null.ok = TRUE, add=argcheck)
  checkmate::assertNumber(v.scale, null.ok=TRUE, add=argcheck)
  checkmate::assertNumber(vertex.label.dist, add=argcheck)
  checkmate::reportAssertions(argcheck)


  if (is.character(plotdat$t1)) {
    if (level=="trt") {
      plotdat <- arrange(plotdat, match(t1, classcode$trt), t2)
    } else if (level=="class") {
      plotdat <- arrange(plotdat, match(t1, classcode$class), t2)
    } else {
      message("Placebo / Pill placebo not included in `plotdat`")
    }
  }

  nodes <- unique(c(plotdat$t1, plotdat$t2))


  g <- igraph::graph.empty()
  g <- g + igraph::vertex(nodes)
  ed <- t(matrix(c(plotdat[["t1"]], plotdat[["t2"]]), ncol = 2))
  ed <- as.vector(ed)
  edges <- igraph::edges(ed, weight = plotdat[["nr"]], arrow.mode=0)
  g <- g + edges


  igraph::E(g)$curved <- FALSE # ensure edges are straight

  if (remove.loops==TRUE) {
    g <- igraph::simplify(g, remove.multiple = FALSE, remove.loops = TRUE)
    plotdat <- plotdat[plotdat$t1!=plotdat$t2,]
  }

  # Get classes
  if (is.character(plotdat$t1)) {
    if (level=="trt") {
      trtnames <- names(igraph::V(g))
      classnames <- classcode$class[match(trtnames, classcode$treat)]
    } else if (level=="class") {
      trtnames <- names(igraph::V(g))
      classnames <- classcode$class[match(trtnames, classcode$class)]
    }
  } else if (is.numeric(plotdat$t1)) {
    if (level=="trt") {
      trtnames <- as.numeric(names(igraph::V(g)))
      classnames <- classcode$class[match(trtnames, classcode$treatcode)]
    } else if (level=="class") {
      trtnames <- as.numeric(names(igraph::V(g)))
      classnames <- classcode$class[match(trtnames, classcode$classcode)]
    }
  }
  igraph::V(g)$names <- trtnames
  igraph::V(g)$classes <- classnames

  # USED IN DEPRESSION WHEN THERE WERE SEPARATE GROUPS
  # if (shortnames==TRUE) {
  #   if (level=="trt") {
  #     groupnames <- classcode$group[match(trtnames, classcode$trtshort)]
  #   } else if (level=="class") {
  #     groupnames <- classcode$group[match(trtnames, classcode$classshort)]
  #   }
  # } else {
  #   if (level=="trt") {
  #     groupnames <- classcode$group[match(trtnames, classcode$trt)]
  #   } else if (level=="class") {
  #     groupnames <- classcode$group[match(trtnames, classcode$class)]
  #   }
  # }
  #
  # igraph::V(g)$group <- as.character(groupnames)

  disconnects <- MBNMAdose:::check.network(g)
  if (v.color=="connect") {
    igraph::V(g)$color <- "SkyBlue2"
    igraph::V(g)$color[which(names(igraph::V(g)) %in% disconnects)] <- "white"
  } else if (v.color=="class") {

    # Get large vector of distinct colours using Rcolorbrewer
    cols <- MBNMAdose:::genmaxcols()
    igraph::V(g)$color <- cols[as.numeric(factor(classnames))]

  } else if (v.color=="group") {

    cols <- RColorBrewer::brewer.pal(8, "Set1")
    igraph::V(g)$color <- cols[as.numeric(factor(V(g)$group))]
  }

  # Add participant numbers
  if (!is.null(v.scale)) {
    if (!(all(c("n1", "n2") %in% names(plotdat)))) {
      stop("`n1` or `n2` not included as a column in dataset. Vertices/nodes will all be scaled to be the same size.")
    }

    ndat <- data.frame("t"=c(plotdat$t1, plotdat$t2), "n"=c(plotdat$n1, plotdat$n2))

    size.vec <- vector()
    for (i in seq_along(igraph::V(g))) {
      size.vec <- append(size.vec, max(ndat$n[ndat$t==names(igraph::V(g))[i]]))
    }
    # Scale size.vec by the max node.size
    size.vec <- size.vec/ (max(size.vec)/20)
    size.vec <- size.vec*v.scale

    igraph::V(g)$size <- size.vec
  }

  # Set order of vertices to be by class
  g <- igraph::permute(g, permutation = Matrix::invPerm(order(classnames)))

  if (!is.null(edge.scale)) {
    igraph::E(g)$width <- igraph::E(g)$weight * edge.scale
  }

  lab.locs <- radian.rescale(x=1:length(igraph::V(g)), direction=-1, start=0)


  if (v.color=="class" & legend==TRUE) {
    par(mfrow=c(1,2))
  }
  if (legend==TRUE) {
    par(mar = c(15, 8, 4, 2)) # Set the margin on all sides to 5
  }

  g <- igraph::set_graph_attr(g, name="layout", value=igraph::layout_in_circle)
  g <- igraph::set_vertex_attr(g, name="label.dist", value=vertex.label.dist)
  g <- igraph::set_vertex_attr(g, name="label.degree", value=lab.locs)
  g <- igraph::set_edge_attr(g, name="color", value="grey20")

  igraph::plot.igraph(g, ...)

  # Add legend
  if (legend==TRUE) {
    if (v.color=="class") {
      plot.new()
      graphics::legend(x="bottomleft", legend=unique(igraph::V(g)$classes), pt.bg=unique(igraph::V(g)$color),
                       pch=21, pt.cex=1.5, cex=0.8)
      par(mfrow=c(1,1))
    } else if (v.color=="group") {
      graphics::legend(x=-2.7, y=-1.25, legend=unique(igraph::V(g)$group), pt.bg=unique(igraph::V(g)$color),
                       pch=21, pt.cex=1.5, cex=0.8)
    }
  }

  return(g)
}










#' Plot forest plot of results vs reference plus numerical results
#'
#' @param outcome A string corresponding to the name of an outcome (as saved in a Results workbook)
#' @param level A string indicating results should be plotted at either the `"treatment"` or `"class"` level
#' @param writepdf A logical object indicating whether a pdf should be written to file
#' @param width Width of plot if `writepdf=TRUE`
#' @param height Height of plot if `writepdf=TRUE`
#' @param exp Logical object indicating whether treatment effects given in Results workbook should be exponentiated or not
#' @param cols A character vector of length 2 indicating the names of the columns for the tabular results
#' @param xlim A numeric vector indicating the ticks on the x axis
#' @param refvalue A number indicating where a vertical dashed line for the reference treatment effect should be drawn
#'
#' @importFrom ggplot2 ggplot_add
#'
#' @export
forestcols <- function(outcome, level="treatment", writepdf=TRUE, width=10, height=12, exp=FALSE, cols=c("OR", "95%CrI"),
                       xlim=c(0.1,0.5,1,5,10,25), refvalue=ifelse(exp==FALSE, 0, 1)) {

  # Run checks
  argcheck <- checkmate::makeAssertCollection()
  checkmate::assertCharacter(outcome, len=1, add=argcheck)
  checkmate::assertChoice(level, choices=c("treatment", "class"), add=argcheck)
  checkmate::assertLogical(writepdf, add=argcheck)
  checkmate::assertLogical(exp, add=argcheck)
  checkmate::assertNumber(width, add=argcheck)
  checkmate::assertNumber(height, add=argcheck)
  checkmate::assertCharacter(cols, len=2, add=argcheck)
  checkmate::assertNumeric(xlim, add=argcheck)
  checkmate::assertNumber(refvalue, add=argcheck)
  checkmate::reportAssertions(argcheck)

  # Load NMA data
  bugsres <- openxlsx::read.xlsx(xlsxFile=paste0("Results/", outcome, ".xlsx"), sheet="Effect Size vs Reference")

  # Choose level
  ind <- which(names(bugsres)=="Class")
  if (level=="treatment") {
    bugsres <- bugsres[,1:(ind-1)]
  } else if (level=="class") {
    bugsres <- bugsres[,ind:ncol(bugsres)]
    bugsres <- bugsres %>% rename(Treatment=Class)
  }
  bugsres <- subset(bugsres, !is.na(Median))

  out.df <- bugsres

  if (exp==TRUE) {
    out.df <- out.df %>% mutate(Median=round(exp(Median),2), Lower95=round(exp(Lower95),2), Upper95=round(exp(Upper95),2),
                                median=Median, val2.5pc=Lower95, val97.5pc=Upper95)
    out.df$NeatOutput <- neatcri(out.df)

    out.df$NeatOutput[1] <- "Reference"
  }

  # Remove point estimate from NeatOutput
  out.df$NeatOutput <- gsub("(.+ )(\\(.+\\))", "\\2", out.df$NeatOutput)

  g <- forestplot(out.df, refvalue = refvalue, xlab=cols[1],
                  ylab=ifelse(level=="treatment", "Treatment", "Class")) +
    theme_bw()

  if (exp==TRUE) {
    g <- g + scale_y_continuous(trans="log", breaks=xlim)
  } else {
    #g <- g + coord_flip(ylim=xlim)
    g <- g + scale_y_continuous(breaks=xlim)
  }

  out.df <- out.df %>% arrange(desc(Treatment)) %>% select(Median, NeatOutput) %>%
    rename(!!cols[2]:=NeatOutput, !!cols[1]:=Median)

  p2 <- gridExtra::tableGrob(out.df,
                             rows=NULL, #gpar.corefill = gpar(fill = "white", col = "white"),
                             theme = gridExtra::ttheme_minimal(base_size=8, padding.h=unit(0.1, "mm"),
                                                               core=list(#fg_params = list(fill = "white"),
                                                                 bg_params = list(fill="white"))
                             )
  )

  separators <- replicate(nrow(p2),
                          grid::segmentsGrob(x0 = unit(0,"npc"), x1 = unit(1,"npc"), y1 = unit(0,"npc"), gp=grid::gpar(lwd=1)),
                          simplify=FALSE)
  p2 <- gtable::gtable_add_grob(p2, grobs = separators, l = 1, r = ncol(p2),
                                #l = seq_len(nrow(p2))
                                t=seq_len(nrow(p2))
  )

  # Set widths/heights to 'fill whatever space I have'
  p2$widths <- unit(rep(1, ncol(p2)), "null")
  p2$heights <- unit(rep(1, nrow(p2)), "null")

  # Get height of plot
  # p3 <- ggplot2::ggplot() +
  #   ggplot2::annotation_custom(p2)
  #p3 <- ggplot2::annotation_custom(p2)

  #gout <- g + p3 + patchwork::plot_layout(nrow = 1)
  #gout <- ggplot2::ggplot_add(object=p3, plot=g, object_name="p3")
  #gout <- gout + patchwork::plot_layout(nrow = 1)

  temp <- ggplot2::ggplot() + ggplot2::annotation_custom(p2)
  gout <- g + temp + patchwork::plot_layout(nrow = 1)

  params <- ggplot_build(gout)
  distance <- params$layout$panel_params[[1]]$y.range[2] - params$layout$panel_params[[1]]$y.range[1]
  distance <- params$layout$panel_params[[1]]$y.range[2] + (distance / nrow(out.df))

  print(distance)
  # p3 <- ggplot2::ggplot() +
  #   ggplot2::annotation_custom(p2, ymax=distance)
  # p3 <- ggplot2::annotation_custom(p2, ymax=distance)

  # Patchwork magic
  #gout <- g + p3 + patchwork::plot_layout(nrow = 1)
  #gout <- ggplot2::ggplot_add(object=p3, plot=g, object_name="p3")
  #gout <- gout + patchwork::plot_layout(nrow = 1)

  temp <- ggplot2::ggplot() + ggplot2::annotation_custom(p2, ymax=distance)
  gout <- g + temp + patchwork::plot_layout(nrow = 1)

  if (writepdf==TRUE) {
    savefile <- paste0("Plots/forest_", level, "_", outcome, ".pdf")
    pdf(savefile, width=width, height=height)
    plot(gout)
    dev.off()

    message(paste0("Plot saved to: ", savefile))
  }

  return(gout)
}
