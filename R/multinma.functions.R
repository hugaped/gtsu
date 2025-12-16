# Functions for creating outputs from multinma
# Author: Hugo Pedder
# Date: 2025-12-15

#' Outputs results from a `multinma` NMA model into an Excel workbook
#'
#' @param nma A `multinma` consistency model of class `"stan_nma"`
#' @param ume A `multinma` UME model of class `"stan_nma"`. Only needs to be specified
#' if user wants to add results for the Unrelated Mean Effects model alongside direct
#' estimates from the NMA.
#' @param outstr A string representing an outcome name. This is used to
#' label the outputted Excel workbook.
#' @param modnam A string representing the name of the model used
#' @param decimals The number of decimal places to which to report numerical results
#' @param trt_ref The name of the reference treatment against which the pooled treatment
#' effects will be plotted and shown in an `Effect size` worksheet.
#' @param lower_better Logical to indicate whether lower scores are ranked as better (`TRUE`)
#' or not (`FALSE`)
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
#' @param pval Numeric threshold for p-value for direct vs indirect estimates. Note that this is
#' an approximation as it assumes the posterior is normally distributed, and the data is correlated
#' ...i.e. it is not a true indicator of direct vs indirect (nodesplitting would be needed for that).
#' It should only be used to highlight comparisons for further investigation.
#' @param ... Arguments to be sent to other `multinma` function
#'
#' @export
multinmatoexcel <- function(nma, ume=NULL,
                            outstr,
                            modnam="RE model",
                            decimals=2, model.scale=NULL,
                            trt_ref=nma$network$treatments[1],
                            lower_better=TRUE,
                            devplot=TRUE, netplot=TRUE, forestplot=TRUE, rankplot=FALSE, scalesd=NULL,
                            pval=0.05,
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


  # --- 2. SETUP & STYLING ---

  # Create directory
  if (!dir.exists("Results")) {
    dir.create("Results", showWarnings = FALSE)
  }

  # Create workbook
  writefile <- paste0("Results/", outstr, ".xlsx")
  path <- "Results/"

  wb <- openxlsx::createWorkbook(title=outstr)


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
  openxlsx::writeData(wb, "Title Page", paste0("Analysis Output: ", outstr), startRow=srow, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Model: ", modnam), startRow=srow+1, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Date Generated: ", Sys.Date()), startRow=srow+2, startCol=1)
  openxlsx::writeData(wb, "Title Page", paste0("Reference Treatment: ", trt_ref), startRow=srow+3, startCol=1)

  title_style <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold", fontName = "Arial")
  openxlsx::addStyle(wb, "Title Page", title_style, rows=srow, cols=1)

  # Add logo
  y <- 0.5
  x <- 0.5*2999/564 # Underlying plot dimensions
  #openxlsx::insertImage(wb, sheet="Title Page", startRow=2, startCol=5, width=4, height=2, dpi=300)
  openxlsx::insertImage(wb, file="man/figures/TSU_24_whitebackground.png",
                       sheet="Title Page", startRow=1, startCol=1, width=x, height=y, dpi=300)


  # --- 4. NODE SUMMARY ---
  openxlsx::addWorksheet(wb, sheetName="Node Output")
  node_sum <- summary(nma)$summary
  #node_sum <- cbind(Parameter = rownames(node_sum), node_sum)
  openxlsx::writeDataTable(wb, sheet="Node Output", node_sum, keepNA = TRUE,
                           tableStyle = "TableStyleMedium2")


  # --- 5. INTERVENTION CODES & COUNTS ---

  # Extract and combine all relevant data types from the network object
  data_list <- list(
    nma$network$agd_arm,
    nma$network$agd_contrast
  )

  # Filter out NULL or empty data frames and combine them
  net_dat <- data_list %>%
    purrr::compact() %>% # Remove NULL elements
    purrr::keep(~nrow(.) > 0) %>% # Keep only non-empty data frames
    dplyr::bind_rows() %>%
    # Ensure only distinct study/treatment rows remain for correct counting
    dplyr::distinct(.study, .trt, .trtclass, .keep_all = TRUE)

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

  # Check for class variable
  class_present <- ".trtclass" %in% names(net_dat)

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
  # if (class_present) {
  #   # Sum participant data grouped by class
  #   class_counts <- trt_counts %>%
  #     dplyr::filter(!is.na(class)) %>% # Only include treatments with a class
  #     dplyr::group_by(class) %>%
  #     dplyr::summarise(
  #       N_Studies_Class = dplyr::n_distinct(.study),
  #       Total_Participants_Class = sum(Total_Participants, na.rm=TRUE),
  #       .groups = 'drop'
  #     )
  #   # NOTE: N_Studies is calculated based on treatments within a class, so it's a proxy
  #   # for the number of studies contributing to that class node.
  # }
  #

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
  openxlsx::writeDataTable(wb, sheet="Data", nma$network$agd_contrast, tableStyle = "TableStyleLight9", startRow = row)
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
  message("Written model fit statistics")



  # --- 9. EFFECT SIZES VS REFERENCE ---
  # Check if ref exists
  if (!trt_ref %in% nma$network$treatments) {
    stop(paste0("trt_ref '", trt_ref, "' is not a treatment within nma$network$treatments"))
  }

  # Get Relative Effects (treatment-level)
  relefs_obj <- multinma::relative_effects(nma, trt_ref=trt_ref)
  relefs <- relefs_obj$summary

  # Handle Transformations (SMD or Log scales)
  # multinma output is usually linear predictor.
  # If model.scale is log, we assume output is log(OR) and needs exp.
  # If scalesd is provided, we assume output is SMD and needs multiplication.

  re_df <- relefs %>% dplyr::as_tibble()

  wbstr <- "Effect sizes reported on analysis scale"
  if (!is.null(scalesd)) {
    cols_to_scale <- c("mean", "sd", "2.5%", "25%", "50%", "75%", "97.5%")
    re_df[cols_to_scale] <- re_df[cols_to_scale] * scalesd

    message(paste0("Effects scaled by SD: ", scalesd))
    wbstr <- paste("Analysis effect sizes rescaled by SD:", scalesd)

  } else if (!is.null(model.scale) && model.scale == "log") {
    cols_to_exp <- c("mean", "2.5%", "25%", "50%", "75%", "97.5%")
    re_df[cols_to_exp] <- exp(re_df[cols_to_exp])
    # SD on natural scale is approximate or not applicable directly

    message("Effects exponentiated (Log to Natural scale)")
    wbstr <- "Effect sizes reported on natural scale"
  }

  # Format Output
  re_out <- re_df %>%
    dplyr::mutate(
      #NeatOutput = neatcri_val(`50%`, `2.5%`, `97.5%`, decimals)
      NeatOutput = neatcri2(., decimals=decimals)
    ) %>%
    dplyr::select(Parameter = parameter, Treatment = .trtb, Median = `50%`, Lower = `2.5%`, Upper = `97.5%`, NeatOutput)

  openxlsx::addWorksheet(wb, sheetName="Effect Size vs Reference")


  openxlsx::writeData(wb, "Effect Size vs Reference", wbstr, startRow=1, startCol=1)
  title_style <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold", fontName = "Arial")
  openxlsx::addStyle(wb, "Effect Size vs Reference", title_style, rows=1, cols=1)

  openxlsx::writeData(wb, sheet="Effect Size vs Reference", re_out, startRow=3)
  apply_std_style(wb, "Effect Size vs Reference", re_out, row=3)

  # Forest Plot
  if (forestplot == TRUE) {
    tryCatch({
      # Use multinma's plot method on the stored object
      # Note: We plot the object *before* manual scaling for the graph to ensure consistency,
      # or we rely on ggplot modification.
      p <- plot(relefs_obj, ref_line = 0, .width=c(0.025, 0.975), ...) +
        ggplot2::theme_bw() +
        ggplot2::labs(title = paste0("Relative Effects vs ", trt_ref))

      print(p) # Print to active device to capture
      openxlsx::insertPlot(wb, sheet="Effect Size vs Reference", startRow=2, startCol=8, width=8, height=6, dpi=300)
    }, error = function(e) warning("Forest plot generation failed: ", e$message))
  }



  # NEED TO ADD CLASS-LEVEL TREATMENT EFFECTS


  # --- 10. RANKS ---
  openxlsx::addWorksheet(wb, "Ranks")

  rks <- multinma::posterior_ranks(nma, lower_better = lower_better)
  rk_sum <- rks$summary

  # Rank Probabilities for Plot
  rk_probs <- multinma::posterior_rank_probs(nma, lower_better = lower_better)

  # Formatting Table
  rk_out <- rk_sum %>%
    dplyr::mutate(NeatOutput = neatcri2(., decimals=decimals),
                  mean=round(mean, decimals)) %>%
    dplyr::select(Treatment = .trt, MeanRank = mean, MedianRank = `50%`, Lower = `2.5%`, Upper = `97.5%`, NeatOutput)

  openxlsx::writeData(wb, "Ranks", rk_out)
  apply_std_style(wb, "Ranks", rk_out)

  # Plot Cumulative Ranks
  if (rankplot==TRUE) {
    g_rank <- plot(rk_probs) +
      ggplot2::theme_bw() +
      ggplot2::theme(legend.position = "bottom")

    print(g_rank)
    openxlsx::insertPlot(wb, sheet="Ranks", startRow=2, startCol=8, width=8, height=6, dpi=300)
  }


  # NEED TO DO THE SAME FOR CLASS RANKINGS



  # --- 11. ALL RELATIVE EFFECTS (CONSISTENCY VS UME) ---

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
  } else if (!is.null(model.scale) && model.scale == "log") {
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

  # Write
  openxlsx::addWorksheet(wb, "Treatment Direct Effects")
  openxlsx::writeData(wb, "Treatment Direct Effects", full_res)
  apply_std_style(wb, "Treatment Direct Effects", full_res)

  # Conditional Formatting for P-Values
  if (!is.null(ume)) {
    pval_col <- which(names(full_res) == "P_Value")
    rows_highlight <- which(full_res$P_Value < pval) + 1 # +1 for header
    if(length(rows_highlight) > 0) {
      openxlsx::addStyle(wb, "Treatment Direct Effects", style_highlight, rows = rows_highlight, cols = pval_col)
    }
  }

  message("Written all direct treatment effects")


  # DO SAME FOR CLASS LEVEL EFFECTS


  openxlsx::saveWorkbook(wb=wb, file=writefile, overwrite=TRUE)

  message("FINISHED!!!!")

}
