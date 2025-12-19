# Functions for creating outputs from multinma
# Author: Hugo Pedder
# Date: 2025-12-15

#' Outputs results from a `multinma` NMA model into an Excel workbook
#'
#' @param nma A `multinma` consistency model of class `"stan_nma"`
#' @param ume A `multinma` UME model of class `"stan_nma"`. Only needs to be specified
#' if user wants to add results for the Unrelated Mean Effects model alongside direct
#' estimates from the NMA.
#' @param outcome A string representing an outcome name. This is used to
#' label the outputted Excel workbook.
#' @param filenam A string representing a file path and name. Must end in `.xlsx`
#' @param modnam A string representing the name of the model used
#' @param decimals The number of decimal places to which to report numerical results
#' @param trt_ref The name of the reference treatment against which the pooled treatment
#' effects will be plotted and shown in an `Effect size` worksheet.
#' @param lower_better Logical to indicate whether lower scores are ranked as better (`TRUE`)
#' or not (`FALSE`)
#' @param treatments_to_rank A character vector of treatment names (must be a subset of those
#' in `nma$network$treatments`) to include in rankings. Allows a decision subset treatments in the
#' network to be ranked.
#' @param classes_to_rank A character vector of class names (must be a subset of those in
#' `nma$network$treatments`) to include in rankings.
#' @param devplot Whether to add a Dev-Dev plot showing the deviance of the UME vs NMA
#' model. To allow this, `addume` must be set to `TRUE` (see requirements for this), and
#' the `dev` node of both NMA and UME models must have been monitored.
#' @param netplot Whether to plot a network plot within a `Network Plots` worksheet
#' @param forestplot Whether to plot forest plots in the `Effect size vs reference` worksheet
#' @param rankplot Whether to plot cumulative rankograms in the `Ranks` worksheet
#' @param scalesd A number indicating the SD to use for back-transforming SMD to a
#' particular measurement scale. If left as `NULL` then no transformation will be
#' performed (i.e. leave as `NULL` if not modelling SMDs)
#' @param eform Indicates whether treatment effect outputs should be exponentiated (`TRUE`) or not
#' (i.e. if model was on log scale).
#' @param pval Numeric threshold for p-value for direct vs indirect estimates. Note that this is
#' an approximation as it assumes the posterior is normally distributed, and the data is correlated
#' ...i.e. it is not a true indicator of direct vs indirect (nodesplitting would be needed for that).
#' It should only be used to highlight comparisons for further investigation.
#' @param addlogo Logical to indicate if GTSU logo should be added
#' @param ... Arguments to be sent to other `multinma` function
#'
#' @export
multinmatoexcel <- function(nma, ume=NULL,
                            filenam,
                            outcome="Outcome",
                            modnam="RE model",
                            decimals=2,
                            eform=FALSE, scalesd=NULL,
                            trt_ref=nma$network$treatments[1],
                            lower_better=TRUE,
                            treatments_to_rank=nma$network$treatments,
                            classes_to_rank=nma$network$classes,
                            devplot=TRUE, netplot=TRUE, forestplot=TRUE, rankplot=FALSE,
                            pval=0.05,
                            addlogo=TRUE,
                            ...) {

  # --- 1. CHECKS ---

  # Check multinma models
  argcheck <- checkmate::makeAssertCollection()
  checkmate::assertClass(nma, "stan_nma", add=argcheck)

  if (!is.null(ume)) {
    checkmate::assertClass(ume, "stan_nma")

    # Check data informing both models are same
    if (!identical(nma$network, ume$network)) {
      stop("`network` object for nma and ume are different\nData sources may differ")
    }
  }
  checkmate::reportAssertions(argcheck)

  if (!is.null(scalesd) & eform==TRUE) {
    warning("Both scalesd and eform=TRUE would not typically be specified")
  }

  if (!grepl("\\.xlsx", filenam)) {
    stop("`filenam` must be a .xlsx file")
  }



  # --- 2. SETUP & STYLING ---

  # Create workbook
  wb <- openxlsx::createWorkbook(title=outcome)

  # Test it can be saved to filenam location
  suppressWarnings(openxlsx::saveWorkbook(wb=wb, file=filenam, overwrite=TRUE))


  # Define Styles
  style_header <- openxlsx::createStyle(
    fontName = "Arial", fontSize = 11, fontColour = "white",
    fgFill = "#4F81BD", textDecoration = "bold",
    halign = "center", valign = "center", border = "Bottom"
  )

  style_subheader <- openxlsx::createStyle(
    fontName = "Arial", fontSize = 10, fontColour = "black",
    fgFill = "#DCE6F1", textDecoration = "bold", border = "Bottom"
  )

  style_general <- openxlsx::createStyle(fontName = "Arial", fontSize = 10)

  style_highlight <- openxlsx::createStyle(fgFill = "#FFEB9C", fontColour = "#9C5700") # Yellow for p-values

  # Helper to apply standard styling to a sheet
  apply_std_style <- function(wb, sheet, df, row=1, col=1) {
    openxlsx::addStyle(wb, sheet, style_header, rows=row, cols=col:(col+ncol(df)-1), gridExpand=TRUE)
    openxlsx::setColWidths(wb, sheet, cols=col:(col+ncol(df)-1), widths = "auto")
  }


  # --- 3. TITLE PAGE ---
  openxlsx::addWorksheet(wb, "Title Page")
  srow <- 4
  openxlsx::writeData(wb, "Title Page", paste0("Analysis Output: ", outcome), startRow=srow, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Model: ", modnam), startRow=srow+1, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Date Generated: ", Sys.Date()), startRow=srow+2, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Reference Treatment: ", trt_ref), startRow=srow+3, startCol=1)

  title_style <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold", fontName = "Arial")
  openxlsx::addStyle(wb, "Title Page", title_style, rows=srow, cols=1)

  # Add logo
  if (addlogo) {
    y <- 0.5
    x <- 0.5*2999/564 # Underlying plot dimensions
    #openxlsx::insertImage(wb, sheet="Title Page", startRow=2, startCol=5, width=4, height=2, dpi=300)
    openxlsx::insertImage(wb, file="man/figures/TSU_24_whitebackground.png",
                          sheet="Title Page", startRow=1, startCol=1, width=x, height=y, dpi=300)
  }


  # --- 4. NODE SUMMARY ---
  openxlsx::addWorksheet(wb, sheetName="Node Output")
  node_sum <- summary(nma)$summary
  #node_sum <- cbind(Parameter = rownames(node_sum), node_sum)
  openxlsx::writeDataTable(wb, sheet="Node Output", node_sum, keepNA = TRUE,
                           tableStyle = "TableStyleMedium2")


  # --- 5. INTERVENTION CODES & COUNTS ---

  # Check for class effects - Note that common class effects is same as treatment model
  class_present <- !is.null(nma$network$classes) && (nma$class_effects=="exchangeable")

  # Extract and combine all relevant data types from the network object
  data_list <- list(
    nma$network$agd_arm,
    nma$network$agd_contrast
  )

  # Check columns to keep distinct
  disvars <- c(".study", ".trt")
  if (class_present) {
    disvars <- append(disvars, ".trtclass")
  }
  disvars <- rlang::quos(disvars)


  # Filter out NULL or empty data frames and combine them
  net_dat <- data_list %>%
    purrr::compact() %>% # Remove NULL elements
    purrr::keep(~nrow(.) > 0) %>% # Keep only non-empty data frames
    dplyr::bind_rows() %>%
    # Ensure only distinct study/treatment rows remain for correct counting
    dplyr::distinct(!!!disvars, .keep_all = TRUE)

  # ----------------------------------------------------------------------
  # 5A. COLUMN CHECKS: Participant Size and Class Variable
  # ----------------------------------------------------------------------

  # Determine which column to use for participant count (prioritise .n, then .sample_size)
  if (".n" %in% names(net_dat)) {
    size_col <- rlang::sym(".n")
  } else if (".sample_size" %in% names(net_dat)) {
    size_col <- rlang::sym(".sample_size")
  } else {
    size_col <- NULL
  }

  # Define the core grouping variables as quosures (expressions)
  # Start with 'treat'
  group_vars <- rlang::quos(treat = .trt)

  # Conditionally add the 'class' grouping variable if it is present in the data
  if (class_present) {
    group_vars <- c(group_vars, rlang::quos(class = .trtclass))
  }

  # ----------------------------------------------------------------------
  # 5B. CALCULATE TREATMENT COUNTS (Studies and Participants)
  # ----------------------------------------------------------------------
  if (!is.null(size_col)) {
    # If a size column is found, perform the standard summarisation
    trt_counts <- net_dat %>%
      # Use !!! to splice the list of grouping expressions into group_by
      dplyr::group_by(!!!group_vars) %>%
      dplyr::summarise(
        N_Studies = dplyr::n_distinct(.study),
        Total_Participants = sum(!!size_col, na.rm=TRUE),
        .groups = 'drop'
      )
  } else {
    # Fallback if no participant count column is available
    trt_counts <- net_dat %>%
      dplyr::group_by(!!!group_vars) %>%
      dplyr::summarise(N_Studies = dplyr::n_distinct(.study), .groups = 'drop')
  }

  # Rename columns
  trt_counts <- trt_counts %>%
    dplyr::rename(Treatment=treat)

  if (class_present) {
    trt_counts <- trt_counts %>%
      dplyr::rename(Class=class)
  }


  # ----------------------------------------------------------------------
  # 5C. CALCULATE CLASS COUNTS (If class variable is present)
  # ----------------------------------------------------------------------
  if (class_present) {
    # Sum participant data grouped by class
    class_counts <- trt_counts %>%
      dplyr::filter(!is.na(Class)) %>% # Only include treatments with a class
      dplyr::group_by(Class) %>%
      dplyr::summarise(
        N_Studies_Class = sum(N_Studies, na.rm=TRUE),
        Total_Participants_Class = sum(Total_Participants, na.rm=TRUE),
        .groups = 'drop'
      )

    trt_counts <- dplyr::left_join(trt_counts, class_counts, by = "Class") %>%
      dplyr::arrange(Class)
  }


  # ----------------------------------------------------------------------
  # 5D. COUNT PAIRWISE COMPARISONS (Treatments and Classes)
  # ----------------------------------------------------------------------

  # --- Treatment Comparisons ---
  pairs_trt_df <- net_dat %>%
    dplyr::select(.study, .trt) %>%
    dplyr::distinct() %>%
    dplyr::inner_join(., ., by = ".study", relationship = "many-to-many") %>%
    dplyr::filter(as.character(.trt.x) < as.character(.trt.y)) %>%
    dplyr::group_by(Treat_1 = .trt.x, Treat_2 = .trt.y) %>%
    dplyr::summarise(N_Studies = dplyr::n_distinct(.study), .groups = 'drop')

  # --- Class Comparisons (If class variable is present) ---
  if (class_present) {
    pairs_class_df <- net_dat %>%
      dplyr::select(.study, .trtclass) %>%
      dplyr::distinct() %>%
      dplyr::inner_join(., ., by = ".study", relationship = "many-to-many") %>%
      dplyr::filter(as.character(.trtclass.x) < as.character(.trtclass.y)) %>%
      dplyr::group_by(Class_1 = .trtclass.x, Class_2 = .trtclass.y) %>%
      dplyr::summarise(N_Studies = dplyr::n_distinct(.study), .groups = 'drop')
  }


  # ----------------------------------------------------------------------
  # 5E. WRITE TO EXCEL
  # ----------------------------------------------------------------------
  # Write Treatment/Class Codes/Counts
  openxlsx::addWorksheet(wb, "Intervention Codes")
  openxlsx::writeData(wb, "Intervention Codes", trt_counts, startRow=1)
  # if (class_present) {
  #   openxlsx::writeData(wb, "Intervention Codes", class_counts, startRow=1, startCol=7)
  # }
  apply_std_style(wb, "Intervention Codes", trt_counts)

  openxlsx::addWorksheet(wb, "N studies per comparison")
  openxlsx::writeData(wb, "N studies per comparison", pairs_trt_df, startRow=1)
  apply_std_style(wb, "N studies per comparison", pairs_trt_df)

  if (class_present) {
    openxlsx::writeData(wb, "N studies per comparison", pairs_class_df, startRow=1, startCol = 5)
    apply_std_style(wb, "N studies per comparison", pairs_class_df, col=5)
  }

  message("Written treatment codes and comparison counts")


  # --- 6. INPUT DATA ---
  openxlsx::addWorksheet(wb, sheetName="Data")

  # Add arm-based
  row <- 1

  # Add contrast-based
  if (!is.null(nma$network$agd_arm)) {
    openxlsx::writeDataTable(wb, sheet="Data", nma$network$agd_arm, tableStyle = "TableStyleLight9", startRow = row)
    row <- row + 3 + nrow(nma$network$agd_arm)
  }
  if (!is.null(nma$network$agd_contrast)) {
    openxlsx::writeDataTable(wb, sheet="Data", nma$network$agd_contrast, tableStyle = "TableStyleLight9", startRow = row)
  }
  message("Written input data")


  # --- 7. NETWORK PLOT ---

  # Write network plots tab with numbers of studies
  # Add number of studies in total, numbers of treatments in total, numbers of classes in total
  sum.df <- matrix(ncol=1, c(dplyr::n_distinct(net_dat$.study),
                             dplyr::n_distinct(net_dat$.trt)))
  nam <- c("N studies", "N treatments")

  if (class_present) {
    sum.df <- rbind(sum.df, dplyr::n_distinct(net_dat$.trtclass))
    nam <- append(nam, "N classes")
  }
  rownames(sum.df) <- nam

  openxlsx::addWorksheet(wb=wb, sheetName="Network Plots")
  openxlsx::writeData(wb=wb, sheet="Network Plots", sum.df,
                      colNames=FALSE, rowNames=TRUE)

  if (netplot == TRUE) {

    # Simple Network Plot
    tryCatch({
      p_net <- plot(nma$network, show_trt_class = class_present, ...) +
        ggplot2::guides(fill = ggplot2::guide_legend("Class", override.aes = list(colour = NA))) +
        ggplot2::labs(title = "Network Geometry")

      print(p_net)
      openxlsx::insertPlot(wb, sheet="Network Plots", startRow=5, startCol=2, width=10, height=8, dpi=300)
      message("Network plot added")
    }, error = function(e) warning("Treatment-level network plot failed"))

    # Class Network Plot
    if (class_present) {
      tryCatch({
        p_net <- plot(nma$network, show_trt_class = class_present, level="class", ...) +
          ggplot2::guides(fill = ggplot2::guide_legend("Class", override.aes = list(colour = NA))) +
          ggplot2::labs(title = "Network Geometry (class level)")

        print(p_net)
        openxlsx::insertPlot(wb, sheet="Network Plots", startRow=5, startCol=15, width=10, height=8, dpi=300)
        message("Network plot added")
      }, error = function(e) warning("Class-level network plot failed"))
    }
  }

  message("Written numbers of studies (network plots tab)")


  # --- 8. MODEL FIT STATISTICS ---
  # Calculate DIC/pD
  fit_stats <- multinma::dic(nma)

  modfit <- data.frame(
    Model = modnam,
    DataPoints = sum(sapply(fit_stats$pointwise, FUN=nrow)),
    ResidualDeviance = round(fit_stats$resdev, decimals),
    pD = round(fit_stats$pd, decimals),
    DIC = round(fit_stats$dic, decimals),
    tau=NA
  )

  # If model is RE
  if (nma$trt_effects=="random") {
    modfit$tau <- neatcri2(filter(node_sum, parameter=="tau"), decimals = decimals)
  }

  if(!is.null(ume)) {
    fit_stats_ume <- multinma::dic(ume)
    modfit_ume <- data.frame(
      Model = "Inconsistency (UME)",
      DataPoints = sum(sapply(fit_stats_ume$pointwise, FUN=nrow)),
      ResidualDeviance = round(fit_stats_ume$resdev, decimals),
      pD = round(fit_stats_ume$pd, decimals),
      DIC = round(fit_stats_ume$dic, decimals),
      tau = NA
    )

    node_sum_ume <- summary(ume)$summary

    # If model is RE
    if (ume$trt_effects=="random") {
      modfit_ume$tau <- neatcri2(filter(node_sum_ume, parameter=="tau"), decimals = decimals)
    }

    modfit <- rbind(modfit, modfit_ume)
  }

  openxlsx::addWorksheet(wb, sheetName="Model fit")
  openxlsx::writeData(wb, sheet="Model fit", modfit)
  apply_std_style(wb, "Model fit", modfit)

  if (devplot==TRUE) {

    if(!is.null(ume)) {
      g_dev <- plot(fit_stats, fit_stats_ume, show_uncertainty = FALSE, ...) +
        ggplot2::theme_bw() +
        ggplot2::xlab("NMA") +
        ggplot2::ylab("UME")

      print(g_dev)
      openxlsx::insertPlot(wb, sheet="Model fit", startRow=5, startCol=1, width=8, height=6, dpi=300)
    } else {
      warning("Can only plot devplot if ume model is provided")
    }
  }

  message("Written model fit statistics")



  # --- 9. EFFECT SIZES VS REFERENCE ---
  # Check if ref exists
  if (!trt_ref %in% nma$network$treatments) {
    stop(paste0("trt_ref '", trt_ref, "' is not a treatment within nma$network$treatments"))
  }

  # 9.1. TREATMENT-LEVEL EFFECTS
  relefs_obj <- multinma::relative_effects(nma, trt_ref=trt_ref)
  relefs <- relefs_obj$summary
  re_df <- relefs %>% dplyr::as_tibble()

  # Scaling/transformation
  wbstr <- "Effect sizes reported on analysis scale"
  if (!is.null(scalesd)) {
    cols_to_scale <- c("mean", "sd", "2.5%", "25%", "50%", "75%", "97.5%")
    re_df[cols_to_scale] <- re_df[cols_to_scale] * scalesd
    wbstr <- paste("Analysis effect sizes rescaled by SD:", scalesd)
  } else if (eform==TRUE) {
    cols_to_exp <- c("mean", "2.5%", "25%", "50%", "75%", "97.5%")
    re_df[cols_to_exp] <- exp(re_df[cols_to_exp])
    wbstr <- "Effect sizes reported on natural scale"
  }

  re_out <- re_df %>%
    dplyr::mutate(NeatOutput = neatcri2(., decimals=decimals)) %>%
    dplyr::select(Parameter = parameter, Treatment = .trtb, Median = `50%`, Lower = `2.5%`, Upper = `97.5%`, NeatOutput)

  # Write Treatment Results
  openxlsx::addWorksheet(wb, sheetName="Effect Size vs Reference")

  refstr <- paste("Effect sizes versus", trt_ref)
  openxlsx::writeData(wb, "Effect Size vs Reference", refstr, startRow=1, startCol=1)
  openxlsx::addStyle(wb, "Effect Size vs Reference", title_style, rows=1, cols=1)

  openxlsx::writeData(wb, "Effect Size vs Reference", wbstr, startRow=2, startCol=1)
  openxlsx::addStyle(wb, "Effect Size vs Reference", title_style, rows=1, cols=1)

  openxlsx::writeData(wb, sheet="Effect Size vs Reference", re_out, startRow=4)
  apply_std_style(wb, "Effect Size vs Reference", re_out, row=4)

  # Insert Treatment Forest Plot (Resized)
  if (forestplot == TRUE) {
    p_trt <- plot(relefs_obj, ref_line = 0, .width=c(0.025, 0.975)) +
      ggplot2::theme_bw() +
      ggplot2::labs(title = paste0("Treatment Effects vs ", trt_ref))

    print(p_trt)
    # Dynamic height
    trt_plot_height <- max(2, (nrow(re_out) * 0.20) + 0.8)
    openxlsx::insertPlot(wb, sheet="Effect Size vs Reference", startRow=2, startCol=8, width=8, height=trt_plot_height, dpi=300)
  }

  # 9.2. CLASS-LEVEL EFFECTS (If present)
  if (class_present==TRUE) {

    # Find class of the reference treatment
    ref_class <- trt_counts$Class[which(trt_counts$Treatment %in% trt_ref)]

    if (ref_class==nma$network$classes[1]) {

      re_class_out <- as_tibble(summary(nma, pars = "class_mean")) %>%
        # Extract class details
        mutate(Class = factor(gsub(".*\\[(.+)\\]", "\\1", parameter),
                              levels = levels(nma$network$classes))) %>%
        mutate(NeatOutput=neatcri2(., decimals=2)) %>%
        dplyr::select(Parameter=parameter, Class, Median=`50%`, Lower=`2.5%`, Upper=`97.5%`, NeatOutput)

    } else {
      # 1. Extract MCMC samples for class means
      # This returns a matrix where rows = iterations, cols = classes
      class_samps <- as.matrix(nma, pars = "class_mean")

      # Add zero column for network reference class
      zmat <- matrix(0, ncol=1, nrow=nrow(class_samps))
      colnames(zmat) <- paste0("class_mean[", nma$network$classes[1], "]")
      class_samps <- cbind(zmat, class_samps)

      # 2. Identify the column corresponding to the Reference Class
      # The column names usually look like "class_mean[ClassName]"
      ref_class <- trt_counts$Class[which(trt_counts$Treatment %in% trt_ref)]
      ref_col_name <- paste0("class_mean[", ref_class, "]")

      if (!ref_col_name %in% colnames(class_samps)) {
        stop(paste("Reference class column", ref_col_name, "not found in MCMC samples."))
      }

      # 3. Calculate Relative Effects (Iteration-by-Iteration)
      # We subtract the reference class column from ALL columns for every iteration.
      # This preserves the correlation structure.
      diff_samps <- sweep(class_samps, 1, class_samps[, ref_col_name], "-")

      # 4. Summarize the new relative samples
      # Apply summary to every column (class)
      class_res_matrix <- t(apply(diff_samps, 2, get_sum))

      # Convert to Tibble and Format
      re_class_out <- as_tibble(class_res_matrix, rownames = "parameter") %>%
        mutate(
          Class = gsub(".*\\[(.+)\\]", "\\1", parameter),
          level = "class"
        ) %>%
        mutate(NeatOutput=neatcri2(., decimals=2)) %>%
        dplyr::select(Parameter=parameter, Class, Median=`50%`, Lower=`2.5%`, Upper=`97.5%`, NeatOutput)
    }

    # 5. Write to Excel
    start_row_class <- nrow(re_out) + 8
    openxlsx::writeData(wb, "Effect Size vs Reference", "Class-level Relative Effects", startRow = start_row_class - 1)
    openxlsx::addStyle(wb, "Effect Size vs Reference", title_style, rows = start_row_class - 1, cols = 1)

    openxlsx::writeData(wb, "Effect Size vs Reference", re_class_out, startRow = start_row_class)
    apply_std_style(wb, "Effect Size vs Reference", re_class_out, row = start_row_class)

    # 6. Insert Class Forest Plot
    if (forestplot == TRUE) {
      # Generate manual plot from summary data
      p_class <- ggplot2::ggplot(re_class_out, ggplot2::aes(y = Class, x = Median)) +
        ggplot2::geom_vline(xintercept = 0, linetype = "dashed", color = "grey50") +
        ggplot2::geom_pointrange(ggplot2::aes(xmin = Lower, xmax = Upper)) +
        ggplot2::theme_bw() +
        ggplot2::labs(
          title = paste("Class Effects relative to", ref_class),
          x = "Relative Effect",
          y = "Class"
        )

      print(p_class)
      # Dynamic height
      class_plot_height <- max(2, (nrow(re_class_out) * 0.20) + 0.8)
      openxlsx::insertPlot(wb, sheet="Effect Size vs Reference", startRow = start_row_class, startCol = 8, width=8, height=class_plot_height, dpi=300)
    }

    message("Written relative effects vs reference class")
  }


  # --- 10. RANKS ---
  openxlsx::addWorksheet(wb, "Ranks")

  # 10.1 Treatment Ranks
  # Create matrix of class effects
  d <- as.matrix(nma, pars = "d")

  zmat <- matrix(0, ncol=1, nrow=nrow(d))
  colnames(zmat) <- paste0("d[", nma$network$treatments[1], "]")
  d <- cbind(zmat, d)

  # Select treatments to rank
  treatments_to_rank <- paste0("class_mean[", treatments_to_rank, "]")
  if (!all(treatments_to_rank %in% colnames(d))) {
    stop("treatments_to_rank do not all match treatment names within nma")
  }
  d <- d[,treatments_to_rank]

  # Calculate ranks at each iteration
  if (lower_better==FALSE) {
    d_rks <- t(apply(-d, 1, rank))
  } else if (lower_better==TRUE) {
    d_rks <- t(apply(d, 1, rank))
  }


  # Calculate Ranks
  d_rk_sum <- summary_multinma_matrix(d_rks)

  # Calculate rank probs
  d_rk_probs <- apply(d_rks, 2, function(x) table(factor(x, levels = 1:ncol(d_rks))) / nrow(d_rks))
  colnames(d_rk_probs) <- gsub("(d\\[)(.+)(\\])", "\\2", colnames(d_rk_probs))
  d_rk_probs <- as.data.frame(d_rk_probs)

  d_rk_probs <- d_rk_probs %>%
    mutate(Rank = row_number()) %>%
    tidyr::pivot_longer(
      cols = -Rank,
      names_to = "Treatment",
      values_to = "Probability"
    )

  # Format Table
  rk_out <- d_rk_sum %>%
    dplyr::mutate(NeatOutput = neatcri2(., decimals=decimals),
                  Class=gsub("(d\\[)(.+)(\\])", "\\2", parameter)) %>%
    dplyr::select(Treatment, MeanRank = mean, MedianRank = `50%`, Lower = `2.5%`, Upper = `97.5%`, NeatOutput)


  openxlsx::writeData(wb, "Ranks", "Treatment Rankings", startRow = 1)
  openxlsx::addStyle(wb, "Ranks", title_style, rows=1, cols=1)

  openxlsx::writeData(wb, "Ranks", rk_out, startRow = 2)
  apply_std_style(wb, "Ranks", rk_out, row=2)

  # Plot Cumulative Ranks (Treatments)
  if (rankplot==TRUE) {
    g_rank <- ggplot2::ggplot(d_rk_probs, ggplot2::aes(x = Rank, y = Probability)) +
      ggplot2::geom_line() +
      ggplot2::facet_wrap(~ Treatment) +
      ggplot2::theme_bw() +
      ggplot2::labs(title = "Cumulative Rank Probabilities (Treatments)",
                    x = "Rank",
                    y = "Probability")

    print(g_rank)
    openxlsx::insertPlot(wb, sheet="Ranks", startRow=2, startCol=8, width=8, height=6, dpi=300)
  }


  # 10.2 Class Ranks
  if (class_present) {

    # Create matrix of class effects
    class_sum <- as.matrix(nma, pars = "class_mean")

    zmat <- matrix(0, ncol=1, nrow=nrow(class_sum))
    colnames(zmat) <- paste0("class_mean[", nma$network$classes[1], "]")
    class_sum <- cbind(zmat, class_sum)
    class_rank_sum <- class_sum

    # Select classes to rank
    classes_to_rank <- paste0("class_mean[", classes_to_rank, "]")
    if (!all(classes_to_rank %in% colnames(class_rank_sum))) {
      stop("classes_to_rank do not all match class names within nma")
    }
    class_rank_sum <- class_rank_sum[,classes_to_rank]

    # Calculate ranks at each iteration
    if (lower_better==FALSE) {
      class_rks <- t(apply(-class_rank_sum, 1, rank))
    } else if (lower_better==TRUE) {
      class_rks <- t(apply(class_rank_sum, 1, rank))
    }


    # Calculate Ranks
    class_rk_sum <- summary_multinma_matrix(class_rks)

    # Calculate rank probs for class
    class_rk_probs <- apply(class_rks, 2, function(x) table(factor(x, levels = 1:ncol(class_rks))) / nrow(class_rks))
    colnames(class_rk_probs) <- gsub("(class_mean\\[)(.+)(\\])", "\\2", colnames(class_rk_probs))
    class_rk_probs <- as.data.frame(class_rk_probs)

    class_rk_probs <- class_rk_probs %>%
      mutate(Rank = row_number()) %>%
      tidyr::pivot_longer(
        cols = -Rank,
        names_to = "Class",
        values_to = "Probability"
      )

    # Format Table
    class_rk_out <- class_rk_sum %>%
      dplyr::mutate(NeatOutput = neatcri2(., decimals=decimals),
                    Class=gsub("(class_mean\\[)(.+)(\\])", "\\2", parameter)) %>%
      dplyr::select(Class, MeanRank = mean, MedianRank = `50%`, Lower = `2.5%`, Upper = `97.5%`, NeatOutput)

    # Determine Row to write to (leave a gap after treatment table)
    start_row_class <- nrow(rk_out) + 5

    openxlsx::writeData(wb, "Ranks", "Class Rankings", startRow = start_row_class - 1)
    openxlsx::addStyle(wb, "Ranks", title_style, rows=start_row_class - 1, cols=1)

    openxlsx::writeData(wb, "Ranks", class_rk_out, startRow = start_row_class)
    apply_std_style(wb, "Ranks", class_rk_out, row=start_row_class)

    # Plot Cumulative Ranks (Classes)
    if (rankplot == TRUE) {

      g_rank_class <- ggplot2::ggplot(class_rk_probs, ggplot2::aes(x = Rank, y = Probability)) +
        ggplot2::geom_line() +
        ggplot2::facet_wrap(~ Class) +
        ggplot2::theme_bw() +
        ggplot2::labs(title = "Cumulative Rank Probabilities (Classes)",
                      x = "Rank",
                      y = "Probability")

      print(g_rank_class)
      # Place plot below the treatment plot.
      # Treatment plot is at row 2, height 6 inches (~30 rows).
      # A safe row estimation for the next plot is around row 35.
      openxlsx::insertPlot(wb, sheet="Ranks", startRow=start_row_class, startCol=8, width=8, height=6, dpi=300)
    }

    message("Written class rankings")
  }


  #######################################################################
  # --- 11. ALL TREATMENT-LEVEL RELATIVE EFFECTS (CONSISTENCY VS UME) ---
  #######################################################################

  # Get all contrasts from NMA
  all_re <- multinma::relative_effects(nma, all_contrasts = TRUE)
  nma_sum <- all_re$summary %>%
    dplyr::as_tibble() %>%
    dplyr::mutate(comp_id = paste0(.trtb, "_vs_", .trta))

  # Prepare main output table
  full_res <- nma_sum %>%
    dplyr::select(Treat1 = .trta, Treat2 = .trtb, comp_id,
                  Median_NMA = `50%`, Low_NMA = `2.5%`, High_NMA = `97.5%`, SD_NMA = sd)

  # UME Comparison logic
  if (!is.null(ume)) {
    # Extract UME summaries.
    # Note: UME models in multinma produce estimates for designs.
    # We essentially want the summary of the split nodes.
    # However, relative_effects(ume) calculates the implied contrasts.
    # In a UME model, these *are* the direct effects (mostly).

    ume_sum <- node_sum_ume %>%
      dplyr::filter(grepl("^d\\[", parameter)) %>%
      dplyr::mutate(comp_id = gsub("(^d\\[)(.+)( vs\\. )(.+)(\\])",
                                   "\\2_vs_\\4",
                                   parameter)) %>%
      dplyr::select(comp_id, Median_UME = `50%`, Low_UME = `2.5%`, High_UME = `97.5%`, SD_UME = sd)

    # Join
    full_res <- dplyr::left_join(full_res, ume_sum, by = "comp_id")

    # Calculate difference (z-score approach for inconsistency)
    # Z = (Mean_NMA - Mean_UME) / sqrt(Var_NMA + Var_UME)
    # Using posterior SD as proxy for SE of estimate
    full_res <- full_res %>%
      dplyr::mutate(
        Diff = Median_NMA - Median_UME,
        SE_Diff = sqrt(SD_NMA^2 + SD_UME^2),
        Z_Score = abs(Diff / SE_Diff),
        P_Value = 2 * (1 - pnorm(Z_Score))
      )
  }

  # Clean up columns and Scale
  if (!is.null(scalesd)) {
    cols <- c("Median_NMA", "Low_NMA", "High_NMA", "Median_UME", "Low_UME", "High_UME", "Diff")
    cols <- intersect(cols, names(full_res))
    full_res[cols] <- full_res[cols] * scalesd

    message(paste("Effects between all treatment comparisons rescaled by:", scalesd))
  } else if (eform==TRUE) {
    cols <- c("Median_NMA", "Low_NMA", "High_NMA", "Median_UME", "Low_UME", "High_UME")
    cols <- intersect(cols, names(full_res))
    full_res[cols] <- exp(full_res[cols])
    # Note: P-value calculation was done on linear scale (correct), transforming results now for display

    message("Effects between all treatment comparisons transformed from log to natural scale")
  }

  # Create Neat Outputs
  full_res$NMA_Result <- neatcri2(full_res, colnams = c("Median_NMA", "Low_NMA", "High_NMA"), decimals)

  if(!is.null(ume)) {
    full_res <- full_res %>%
      mutate(Direct_Result=case_when(!is.na(Median_UME) ~ neatcri2(full_res, colnams = c("Median_UME", "Low_UME", "High_UME"), decimals),
                                            TRUE ~ "")) %>%
      dplyr::select(-comp_id, -SD_NMA, -SD_UME, -SE_Diff, -Z_Score)
  } else {
    full_res <- full_res %>% dplyr::select(-comp_id, -SD_NMA)
  }


  # LEAGUE TABLE WITH NMA (Bottom-Left) AND UME (Top-Right)

  # 1. Get unique treatments in the order they appear in the network
  trts <- as.character(nma$network$treatments)
  n_trts <- length(trts)

  # 2. Initialize an empty square matrix (plus space for labels)
  # We make it (n_trts + 1) x (n_trts + 1) to accommodate the labels
  lt_matrix <- matrix("", nrow = n_trts + 1, ncol = n_trts + 1)

  # 3. Add Labels to the 1st row and 1st column
  lt_matrix[1, 2:(n_trts+1)] <- trts
  lt_matrix[2:(n_trts+1), 1] <- trts
  lt_matrix[1, 1] <- "Ref: Row vs Col"

  # 4. Fill the matrix
  for (i in 1:n_trts) {
    for (j in 1:n_trts) {
      if (i == j) {
        lt_matrix[i+1, j+1] <- "--" # Diagonal
        next
      }

      # Identify the comparison
      t_row <- trts[i]
      t_col <- trts[j]

      if (i > j) {
        # LOWER TRIANGLE: NMA Results (Row vs Column)
        # Find result in full_res (Treat2 is .trtb/Row, Treat1 is .trta/Col)
        res <- full_res %>% dplyr::filter((Treat2 == t_row & Treat1 == t_col))

        if (nrow(res) > 0) {
          lt_matrix[i+1, j+1] <- res$NMA_Result[1]
        } else {
          # Try inverted contrast
          res_inv <- full_res %>% dplyr::filter((Treat1 == t_row & Treat2 == t_col))
          if (nrow(res_inv) > 0) {
            # Need to invert the numeric values before formatting for the neatcri
            # This assumes neatcri2 is available to handle a temporary inverted df
            inv_df <- res_inv %>% mutate(
              m = if(eform==TRUE) 1/Median_NMA else -Median_NMA,
              l = if(eform==TRUE) 1/High_NMA else -High_NMA,
              h = if(eform==TRUE) 1/Low_NMA else -Low_NMA
            )
            lt_matrix[i+1, j+1] <- neatcri2(inv_df, colnams = c("m", "l", "h"), decimals=decimals)
          }
        }

      } else {
        # UPPER TRIANGLE: UME/Direct Results (Row vs Column)
        if (!is.null(ume)) {
          res <- full_res %>% dplyr::filter((Treat2 == t_row & Treat1 == t_col))

          if (nrow(res) > 0 && !is.na(res$Median_UME[1])) {
            lt_matrix[i+1, j+1] <- res$Direct_Result[1]
          } else {
            # Try inverted contrast
            res_inv <- full_res %>% dplyr::filter((Treat1 == t_row & Treat2 == t_col))
            if (nrow(res_inv) > 0 && !is.na(res_inv$Median_UME[1])) {
              # inv_df <- res_inv %>% mutate(
              #   m = if(eform==TRUE) 1/Median_UME else -Median_UME,
              #   l = if(eform==TRUE) 1/High_UME else -High_UME,
              #   h = if(eform==TRUE) 1/Low_UME else -Low_UME
              # )
              # lt_matrix[i+1, j+1] <- neatcri2(inv_df, colnams = c("m", "l", "h"), decimals=decimals)
              lt_matrix[i+1, j+1] <- neatcri2(res_inv, colnams = c("Median_UME", "Low_UME", "High_UME"), decimals=decimals)
            } else {
              lt_matrix[i+1, j+1] <- ""
            }
          }
        }
      }
    }
  }

  # --- WRITE TO EXCEL ---
  openxlsx::addWorksheet(wb, "Treatment Direct Effects")
  openxlsx::writeData(wb, "Treatment Direct Effects", full_res)
  apply_std_style(wb, "Treatment Direct Effects", full_res)

  # Write League Table
  # Positioned to the right of the main results table
  start_col_league <- ncol(full_res) + 3
  league_table_df <- as.data.frame(lt_matrix)

  # Write a descriptive title above the league table
  openxlsx::writeData(wb, "Treatment Direct Effects", "League Table: Lower=NMA (Row vs Col, Green), Upper=Direct (Col vs Row, Grey)",
                      startCol = start_col_league, startRow = 1)

  # Write the matrix without default data frame headers
  openxlsx::writeData(wb, "Treatment Direct Effects", league_table_df,
                      startCol = start_col_league, startRow = 2, colNames = FALSE)

  # ----------------------------------------------------------------------
  # STYLING: COLUMN WIDTHS AND TREATMENT LABELS
  # ----------------------------------------------------------------------

  # Define specific coordinates
  n_trts <- length(trts)
  league_rows <- 2:(n_trts + 2) # Includes header row
  league_cols <- start_col_league:(start_col_league + n_trts)

  # 1. Increase column widths for readability
  # We set a fixed width (e.g., 20) or use a multiplier based on decimal precision
  # Column 1 (labels) can be slightly narrower or wider depending on trt name length
  openxlsx::setColWidths(wb, "Treatment Direct Effects",
                         cols = league_cols,
                         widths = 16)

  # 2. Create and Apply Specific Styles
  style_nma_bg <- openxlsx::createStyle(fgFill = "#EBF1DE", border = "TopBottomLeftRight", halign = "center", valign = "center")
  style_ume_bg <- openxlsx::createStyle(fgFill = "#F2F2F2", border = "TopBottomLeftRight", halign = "center", valign = "center")
  style_diag   <- openxlsx::createStyle(fgFill = "#D9D9D9", border = "TopBottomLeftRight", halign = "center", valign = "center")

  # 3. Apply Triangle and Diagonal Styling
  for (i in 1:n_trts) {
    for (j in 1:n_trts) {
      target_row <- i + 2 # Offset for title row
      target_col <- start_col_league + j # Offset for label column

      if (i == j) {
        # Diagonal cells
        openxlsx::addStyle(wb, "Treatment Direct Effects", style_diag, rows = target_row, cols = target_col)
      } else if (i > j) {
        # Lower Triangle (NMA)
        openxlsx::addStyle(wb, "Treatment Direct Effects", style_nma_bg, rows = target_row, cols = target_col)
      } else {
        # Upper Triangle (UME/Direct)
        openxlsx::addStyle(wb, "Treatment Direct Effects", style_ume_bg, rows = target_row, cols = target_col)
      }
    }
  }

  # 4. Format Treatment Names (Row 1 and Col 1 of the League Table)
  # Using the standard header style defined at the top of the function
  openxlsx::addStyle(wb, "Treatment Direct Effects", style_header,
                     rows = 2, cols = league_cols, gridExpand = TRUE)
  openxlsx::addStyle(wb, "Treatment Direct Effects", style_header,
                     rows = league_rows, cols = start_col_league, gridExpand = TRUE)

  # Ensure the "Ref" cell at the top-left of the table is also styled
  openxlsx::addStyle(wb, "Treatment Direct Effects", style_header, rows = 2, cols = start_col_league)

  message("Written all direct treatment effects with league table")


  #######################################################################
    # --- 12. ALL Class-LEVEL RELATIVE EFFECTS (CONSISTENCY ONLY) ---
  #######################################################################

  if (class_present) {

    # Get all contrast pairs
    class_nams <- gsub("(class_mean\\[)(.+)(\\])", "\\2", colnames(class_sum))
    colnames(class_sum) <- class_nams
    contrast_pairs <- combn(class_nams, 2, simplify = FALSE)

    # Iterate through pairs and calculate the relative effect samples
    # We calculate: Class B - Class A
    all_contrasts_list <- lapply(contrast_pairs, function(pair) {
      trta <- pair[1] # Reference
      trtb <- pair[2] # Target

      # Calculate iteration-by-iteration difference
      diff_samples <- class_sum[, trtb] - class_sum[, trta]

      # Summarize the posterior samples
      # (Assumes get_sum or a similar summary function is available)
      res_class <- data.frame(
        .trta = trta,
        .trtb = trtb,
        mean = mean(diff_samples),
        sd = sd(diff_samples),
        `2.5%` = quantile(diff_samples, 0.025),
        `50%` = quantile(diff_samples, 0.5),
        `97.5%` = quantile(diff_samples, 0.975),
        check.names = FALSE
      )
      return(res_class)
    })


    # Combine into a single data frame
    full_res_class <- do.call(rbind, all_contrasts_list) %>%
      as_tibble() %>%
      dplyr::select(Class1=.trta, Class2=.trtb, Median_NMA=`50%`, Low_NMA=`2.5%`, High_NMA=`97.5%`)

    # Clean up columns and Scale
    if (!is.null(scalesd)) {
      cols <- c("Median_NMA", "Low_NMA", "High_NMA")
      cols <- intersect(cols, names(full_res_class))
      full_res_class[cols] <- full_res_class[cols] * scalesd

      message(paste("Effects between all treatment comparisons rescaled by:", scalesd))
    } else if (eform==TRUE) {
      cols <- c("Median_NMA", "Low_NMA", "High_NMA")
      cols <- intersect(cols, names(full_res_class))
      full_res_class[cols] <- exp(full_res_class[cols])

      message("Effects between all treatment comparisons transformed from log to natural scale")
    }

    # Create Neat Outputs
    full_res_class$NMA_Result <- neatcri2(full_res_class, colnams = c("Median_NMA", "Low_NMA", "High_NMA"), decimals)


    # CLASS-LEVEL LEAGUE TABLE WITH NMA (Bottom-Left)

    n_class <- length(class_nams)

    # Initialize an empty square matrix (plus space for labels)
    lt_matrix_class <- matrix("", nrow = n_class + 1, ncol = n_class + 1)

    # Add Labels to the 1st row and 1st column
    lt_matrix_class[1, 2:(n_class+1)] <- class_nams
    lt_matrix_class[2:(n_class+1), 1] <- class_nams
    lt_matrix_class[1, 1] <- "Ref: Row vs Col"

    # Fill the matrix
    for (i in 1:n_class) {
      for (j in 1:n_class) {
        if (i == j) {
          lt_matrix_class[i+1, j+1] <- "--" # Diagonal
          next
        }

        # Identify the comparison
        t_row <- class_nams[i]
        t_col <- class_nams[j]

        if (i > j) {
          # LOWER TRIANGLE: NMA Results (Row vs Column)
          # Find result in full_res_class
          res <- full_res_class %>% dplyr::filter((Class2 == t_row & Class1 == t_col))

          if (nrow(res) > 0) {
            lt_matrix_class[i+1, j+1] <- res$NMA_Result[1]
          } else {
            # Try inverted contrast
            res_inv <- full_res_class %>% dplyr::filter((Class1 == t_row & Class2 == t_col))
            if (nrow(res_inv) > 0) {
              # Need to invert the numeric values before formatting for the neatcri
              # This assumes neatcri2 is available to handle a temporary inverted df
              inv_df <- res_inv %>% mutate(
                m = if(eform==TRUE) 1/Median_NMA else -Median_NMA,
                l = if(eform==TRUE) 1/High_NMA else -High_NMA,
                h = if(eform==TRUE) 1/Low_NMA else -Low_NMA
              )
              lt_matrix_class[i+1, j+1] <- neatcri2(inv_df, colnams = c("m", "l", "h"), decimals=decimals)
            }
          }

        }
      }
    }

    # --- WRITE TO EXCEL ---
    openxlsx::addWorksheet(wb, "Class Direct Effects")
    openxlsx::writeData(wb, "Class Direct Effects", full_res_class)
    apply_std_style(wb, "Class Direct Effects", full_res_class)

    # Write League Table
    # Positioned to the right of the main results table
    start_col_league <- ncol(full_res_class) + 3
    league_table_df <- as.data.frame(lt_matrix_class)

    # Write a descriptive title above the league table
    openxlsx::writeData(wb, "Class Direct Effects", "League Table: Lower=NMA (Row vs Col, Green)",
                        startCol = start_col_league, startRow = 1)

    # Write the matrix without default data frame headers
    openxlsx::writeData(wb, "Class Direct Effects", league_table_df,
                        startCol = start_col_league, startRow = 2, colNames = FALSE)

    # ----------------------------------------------------------------------
    # STYLING: COLUMN WIDTHS AND TREATMENT LABELS
    # ----------------------------------------------------------------------

    # Define specific coordinates
    league_rows <- 2:(n_class + 2) # Includes header row
    league_cols <- start_col_league:(start_col_league + n_class)

    # 1. Increase column widths for readability
    # We set a fixed width (e.g., 20) or use a multiplier based on decimal precision
    # Column 1 (labels) can be slightly narrower or wider depending on trt name length
    openxlsx::setColWidths(wb, "Class Direct Effects",
                           cols = league_cols,
                           widths = 16)

    # Apply Triangle and Diagonal Styling
    for (i in 1:n_class) {
      for (j in 1:n_class) {
        target_row <- i + 2 # Offset for title row
        target_col <- start_col_league + j # Offset for label column

        if (i == j) {
          # Diagonal cells
          openxlsx::addStyle(wb, "Class Direct Effects", style_diag, rows = target_row, cols = target_col)
        } else if (i > j) {
          # Lower Triangle (NMA)
          openxlsx::addStyle(wb, "Class Direct Effects", style_nma_bg, rows = target_row, cols = target_col)
        } else {
          # Upper Triangle (UME/Direct)
          openxlsx::addStyle(wb, "Class Direct Effects", style_ume_bg, rows = target_row, cols = target_col)
        }
      }
    }

    # Format Treatment Names (Row 1 and Col 1 of the League Table)
    # Using the standard header style defined at the top of the function
    openxlsx::addStyle(wb, "Class Direct Effects", style_header,
                       rows = 2, cols = league_cols, gridExpand = TRUE)
    openxlsx::addStyle(wb, "Class Direct Effects", style_header,
                       rows = league_rows, cols = start_col_league, gridExpand = TRUE)

    # Ensure the "Ref" cell at the top-left of the table is also styled
    openxlsx::addStyle(wb, "Class Direct Effects", style_header, rows = 2, cols = start_col_league)

    message("Written all direct class effects with league table")
  }

  openxlsx::saveWorkbook(wb=wb, file=filenam, overwrite=TRUE)

  message("FINISHED!!!!")

}
