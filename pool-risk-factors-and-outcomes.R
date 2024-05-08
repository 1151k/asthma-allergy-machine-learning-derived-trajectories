# Load packages
packages_all = c("data.table", "forestploter", "MetaUtility", "readxl", "robumeta", "metafor", "ggplot2", "grid", "FDRestimation", "openxlsx")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
input_files <- c(
    'meta-analysis--AD-eczema--mid-onset-late-resolving--risk-factors'
)
subgroup_analysis <- F # rating-no-weak; sex-male; sex-female; risk-general; risk-high (or F if including ALL data)
meta_analysis_output <- data.table(
    variable_group = character(),
    n_studies = numeric(),
    pooled_estimate = numeric(),
    pooled_lower_95ci = numeric(),
    pooled_upper_95ci = numeric(),
    pooled_formatted_results = character(),
    se = numeric()
)
p_values <- data.table(
    variable_group = character(),
    p_value_raw = numeric(),
    p_value_raw_formatted = character(),
    p_value_fdr = numeric()
)
# xlim <- c(0.9, 20)
# ticks_at <- c(0.25, 0.5, 1, 2, 4)



for (input_file in input_files) {
    cat('LOAD FILE\n')
    cat('-------- start\n')
    print(paste('Preparing to load', input_file))
    # Read the Excel file
    input_file_name <- paste0(folder_path, 'output/docx/', input_file, '.xlsx')
    output_file_name <- input_file
    df <- readxl::read_excel(input_file_name)
    print('Finished loading Excel file')
    data.table::setDT(df)
    # set analysis_type to 'risk-factors' if  the input_file ends with 'risk-factors', otherwise set it to 'outcomes'
    analysis_type <- ifelse(grepl('risk-factors$', input_file), 'risk-factors', 'outcomes')
    # capitalize the first letter in the column of a few variables
    df$variable_category <- sub("^(.)", "\\U\\1", df$variable_category, perl = TRUE)
    df$variable_group <- sub("^(.)", "\\U\\1", df$variable_group, perl = TRUE)
    df$variable_value <- sub("^(.)", "\\U\\1", df$variable_value, perl = TRUE)
    df$sex <- sub("^(.)", "\\U\\1", df$sex, perl = TRUE)
    df$overall_rating <- sub("^(.)", "\\U\\1", df$overall_rating, perl = TRUE)
    df$risk_group <- sub("^(.)", "\\U\\1", df$risk_group, perl = TRUE)
    at_least_one_meta_analysis_performed <- F
    # if subgroup analysis is to be done, set things up here
    if (!is.na(subgroup_analysis)) {
        if (subgroup_analysis == 'rating-no-weak') {
            # remove rows where overall_rating is 'Weak'
            df <- df[overall_rating != 'Weak',]
            # if one or less rows, abort program
            if (nrow(df) <= 1) {
                print('Too few data points to run sensitivity analysis (i.e., excluding studies with Low overall rating)')
                quit()
            } else {
                print('Preparing to run sensitivity analysis by excluding studies with Low overall rating')
                output_file_name <- paste0(output_file_name, '--rating-no-weak')
            }
        } else if (subgroup_analysis == 'sex-male') {
            # remove rows where sex is 'Female' or 'Mixed'
            df <- df[sex != 'Female' & sex != 'Mixed',]
            # if one or less rows, abort program
            if (nrow(df) <= 1) {
                print('Too few data points to run subgroup analysis by male sex')
                quit()
            } else {
                print('Preparing to run subgroup analysis by male sex')
                output_file_name <- paste0(output_file_name, '--sex-male')
            }
        } else if (subgroup_analysis == 'sex-female') {
            # remove rows where sex is 'Male' or 'Mixed'
            df <- df[sex != 'Male' & sex != 'Mixed',]
            # if one or less rows, abort program
            if (nrow(df) <= 1) {
                print('Too few data points to run subgroup analysis by female sex')
                quit()
            } else {
                print('Preparing to run subgroup analysis by female sex')
                output_file_name <- paste0(output_file_name, '--sex-female')
            }
        } else if (subgroup_analysis == 'risk-general') {
            # remove rows where risk is 'High'
            df <- df[risk_group != 'High-risk population' & risk_group != 'Mixed',]
            # if one or less rows, abort program
            if (nrow(df) <= 1) {
                print('Too few data points to run subgroup analysis by general risk')
                quit()
            } else {
                print('Preparing to run subgroup analysis by general risk')
                output_file_name <- paste0(output_file_name, '--risk-general')
            }
        } else if (subgroup_analysis == 'risk-high') {
            # remove rows where risk is 'General'
            df <- df[risk_group != 'General population' & risk_group != 'Mixed',]
            # if one or less rows, abort program
            if (nrow(df) <= 1) {
                print('Too few data points to run subgroup analysis by high risk')
                quit()
            } else {
                print('Preparing to run subgroup analysis by high risk')
                output_file_name <- paste0(output_file_name, '--risk-high')
            }
        }
    }
    print('Finished preparing data')
    cat('-------- end\n')
    # run meta-analysis on each subgroup
    cat('RUN META-ANALYSIS BY SUBGROUP\n')
    cat('-------- start\n')
    # loop through each unique value in df$variale_category in alphabetical order
    predefined_ordered_variable_categories <- c(
        # risk factors
        'Sociodemographics',
        'Prenatal exposure / hereditary',
        'Early-life exposure',
        # clinical outcomes
        'Allergic disease',
        'Allergic sensitization',
        'Lung function',
        # other
        'Other'
    )
    present_ordered_variable_categories <- c()
    for (variable_category_value in predefined_ordered_variable_categories) {
        if (variable_category_value %in% unique(df$variable_category)) {
            present_ordered_variable_categories <- c(present_ordered_variable_categories, variable_category_value)
        }
    }
    only_RR <- T
    ever_RR <- F
    at_least_one_meta_analysis_performed_in_category <- F
    for (variable_category_value in present_ordered_variable_categories) {
        at_least_one_meta_analysis_performed_in_category <- F
        unique_variable_values <- unique(df[variable_category == variable_category_value]$variable_group)
        new_rows <- data.table(
            variable_group =  character(),
            n_studies = numeric(),
            pooled_estimate = numeric(),
            pooled_lower_95ci = numeric(),
            pooled_upper_95ci = numeric(),
            pooled_formatted_results = character(),
            se = numeric()
        )
        significant_variables <- 999 # TODO! remove if wanting to only display significant findings
        for (variable_group_value in unique_variable_values) {
            print(paste('Running meta-analysis for', variable_group_value))
            df_sub <- df[variable_group == variable_group_value]
            # check if there are at least 2 studies
            if (length(unique(df_sub$study_id)) < 2) {
                print('Too few data points to run meta-analysis')
                next
            } else {
                at_least_one_meta_analysis_performed <- T
                at_least_one_meta_analysis_performed_in_category <- T
            }
            ####
            ####
            if (any(grepl("RR|RRR", df_sub$outcome_measure))) {
                OR_to_RR <- TRUE
                ever_RR <- TRUE
            } else {
                OR_to_RR <- FALSE
                only_RR <- F
            }
            # # check df$outcome_measure, if it's OR and common == 'yes', set effect_size = sqrt(effect_size) and upper_95ci = sqrt(upper_95ci)
            if (OR_to_RR) {
                print('Transforming OR to RR where common == yes')
                df_sub[, point_estimate := ifelse(outcome_measure == 'OR' & common == 'yes', sqrt(point_estimate), point_estimate)]
                df_sub[, upper_95ci := ifelse(outcome_measure == 'OR' & common == 'yes', sqrt(upper_95ci), upper_95ci)]
                df_sub[, lower_95ci := ifelse(outcome_measure == 'OR' & common == 'yes', sqrt(lower_95ci), lower_95ci)]
            }
            if (any(grepl("MD|SMD", df_sub$outcome_measure))) {
                ma_data <- MetaUtility::scrape_meta(type = "raw", df_sub$point_estimate, df_sub$upper_95ci)
                df_sub$log_estimate_point_estimate <- df_sub$point_estimate
                df_sub$log_estimate_variance <- ma_data$vyi
                logarithmic_scale <- F
            } else {
                ma_data <- MetaUtility::scrape_meta(type = "RR", df_sub$point_estimate, df_sub$upper_95ci)
                df_sub$log_estimate_point_estimate <- ma_data$yi
                df_sub$log_estimate_variance <- ma_data$vyi
                logarithmic_scale <- T
            }
            ####
            ####
            n_data <- length(unique(df_sub$data_id))
            n_studies <- length(unique(df_sub$study_id))
            dfs_too_low <- FALSE
            if (n_studies < n_data) {
                print('At least one study with multiple data points')
                run_rve = TRUE
                # test RVE
                meta_analysis <- robumeta::robu(
                    log_estimate_point_estimate ~ 1,
                    data = df_sub,
                    studynum = study_id,
                    modelweights = "CORR",
                    var.eff.size = log_estimate_variance,
                    small = TRUE,
                    rho = 0.8
                )
                if (meta_analysis$dfs < 4) {
                    dfs_too_low <- TRUE
                    print('Too low degrees of freedom for RVE')
                    df_text <- paste0('RVE unreliable (df=', round(meta_analysis$dfs,2), '), ')
                    run_rve <- FALSE
                }
            } else {
                run_rve <- FALSE
            }
            if (run_rve) {
                effect_size <- meta_analysis$b.r
                lower_95ci <- meta_analysis$reg_table$CI.L
                upper_95ci <- meta_analysis$reg_table$CI.U
                se <- meta_analysis$reg_table$SE
                i_squared <- meta_analysis$mod_info$I.2
                tau_squared <- meta_analysis$mod_info$tau.sq
                dfs <- meta_analysis$dfs
                p <- meta_analysis$reg_table$prob
                df_text <- paste0('df=', dfs, ', ')
            } else {
                meta_analysis <- metafor::rma.uni(
                    yi = log_estimate_point_estimate,
                    vi = log_estimate_variance,
                    data = df_sub
                )
                # print('Naive meta-analysis')
                # print(summary(meta_analysis))
                effect_size <- meta_analysis$beta
                lower_95ci <- meta_analysis$ci.lb
                upper_95ci <- meta_analysis$ci.ub
                se <- meta_analysis$se
                i_squared <- meta_analysis$I2
                tau_squared <- meta_analysis$tau2
                p <- meta_analysis$pval
                if (!dfs_too_low) {
                    df_text <- ''
                }
            }
            # set p-value
            p_values <- rbind(
                p_values,
                data.table(
                    variable_group = variable_group_value,
                    p_value_raw = p,
                    p_value_raw_formatted = format.pval(p, digits = 4),
                    p_value_fdr = NA
                ),
                use.names = FALSE
            )
            # some general processing of results
            if (logarithmic_scale) {
                effect_size <- round(metafor::transf.exp.int(effect_size), 2)
                lower_95ci <- round(metafor::transf.exp.int(lower_95ci), 2)
                upper_95ci <- round(metafor::transf.exp.int(upper_95ci), 2)
            } else {
                effect_size <- round(effect_size, 2)
                lower_95ci <- round(lower_95ci, 2)
                upper_95ci <- round(upper_95ci, 2)
            }
            formatted_effect <- paste0(effect_size, " (", lower_95ci, ", ", upper_95ci, ")")
            i_squared <- round(i_squared, 1)
            tau_squared <- round(tau_squared, 2)
            if (p < 0.01) {
                # p <- format.pval(p, digits = 2, scientific = TRUE)
                p <- '<0.01'
            } else {
                p <- paste0('=',round(p, 2))
            }
            # print(meta_analysis)
            # check if significant 95%CI
            if (lower_95ci > 1 | upper_95ci < 1) {
                significant_variables <- significant_variables + 1
            }
            # add to the output table
            new_row <- data.table(
                variable_group = paste0(variable_group_value, ifelse(OR_to_RR, '*', '')),
                n_studies = n_studies,
                pooled_estimate = effect_size,
                pooled_lower_95ci = lower_95ci,
                pooled_upper_95ci = upper_95ci,
                pooled_formatted_results = formatted_effect,
                se = se
            )
            new_rows <- rbind(new_rows, new_row, use.names = FALSE)
            # make a forest plot based on the results in this loop (i.e., results from meta_analysis and data of df_sub)
            df_sub$formatted_results <- paste0(round(df_sub$point_estimate,2), " (", round(df_sub$lower_95ci,2), ", ", round(df_sub$upper_95ci,2), ")")
            df_sub$empty_column <- paste(rep(" ", 33), collapse = " ")
            # sort df_sub by decreasing OR
            df_sub <- df_sub[order(-df_sub$point_estimate)]
            print('df_sub')
            print(df_sub)
            # add a row to df_sub with the pooled results
            df_sub <- rbind(
                df_sub,
                data.table(
                    study_id = "",
                    data_id = "Pooled",
                    variable_category = "",
                    variable_group = "",
                    variable_value = "",
                    outcome_measure = "",
                    common = "",
                    point_estimate = effect_size,
                    lower_95ci = lower_95ci,
                    upper_95ci = upper_95ci,
                    sex = "",
                    risk_group = "",
                    overall_rating = "",
                    log_estimate_point_estimate = "",
                    log_estimate_variance = "",
                    formatted_results = formatted_effect,
                    empty_column = ""
                ),
                use.names = FALSE
            )
            df_sub <- df_sub[, .(data_id, variable_group, variable_value, point_estimate, lower_95ci, upper_95ci, sex, risk_group, overall_rating, formatted_results, empty_column)]
            colname_measure_of_effect <- ifelse(logarithmic_scale, ifelse(OR_to_RR, ifelse(any(grepl("OR", df_sub$outcome_measure)), "OR or RR", "RR"), "OR"), "MD")
            colnames(df_sub) <- c("Data ID", "Variable", "Variable value", "OR", "Lower 95%CI", "Upper 95%CI", "Sex", "Risk group", "Overall rating", paste0(colname_measure_of_effect, ' (95%CI)'), "")
            #
            if (logarithmic_scale) {
                x_lower_lim <- max(0.01, min(df_sub$`Lower 95%CI`, na.rm=T) - 0.1)
                x_lower_lim <- ifelse(x_lower_lim > 1, 0.9, x_lower_lim)
                xlim <- c(
                    x_lower_lim,
                    min(max(df_sub$`OR`, na.rm=T) + 2, max(df_sub$`Upper 95%CI`, na.rm=T))
                )
                geom_mean <- sqrt(min(xlim) * max(xlim))
                p10_val <- quantile(xlim, 0.3)
                geom_p10 <- sqrt(min(xlim) * p10_val)
                p90_val <- quantile(xlim, 0.7)
                geom_p90 <- sqrt(max(xlim) * p90_val)
                ticks_at <- c(geom_p10, geom_mean, geom_p90)
                ticks_at <- round(ticks_at, 1)
                # remove duplicat values in ticks_at
                ticks_at <- unique(ticks_at)
                # in case thre are any ticks within ±0.2 of 1, remove them and add one directly on 1, but if there are ≥2 such, don't do anything
                print('ticks')
                print(ticks_at)
                if (sum(abs(ticks_at - 1) <= 0.2) >= 1) {
                    ticks_at <- ticks_at[abs(ticks_at - 1) >= 0.2]
                    ticks_at <- c(ticks_at, 1)
                } else {
                    ticks_at <- c(ticks_at, 1)
                }
                # order values in ticks_at
                ticks_at <- sort(ticks_at)
                # if x_lim[1] is higher than 1, or x_lim[2] is lower than 1, set corresponding x_lim value to 1
                if (xlim[1] > 1) {
                    xlim[1] <- 0.98
                }
                if (xlim[2] < 1) {
                    xlim[2] <- 1.02
                }
            }
            # set up stuff for raw forest plots
            raw_theme <- forestploter::forest_theme(
                core = list(bg_params=list(fill = c("transparent"))),
                xaxis_lwd = 1.3,
                refline_lwd = 1.3,
                refline_lty = 1,
                base_size = 9,
                footnote_cex = 0.75,
                summary_fill = "#575e92",
                summary_col = "#575e92"
            )
            raw_forest_plot <- forestploter::forest(
                df_sub[, c(1, 3, 7, 8, 9, 10, 11)],
                est = df_sub$`OR`,
                lower = df_sub$`Lower 95%CI`,
                upper = df_sub$`Upper 95%CI`,
                is_summary = c(rep(FALSE, nrow(df_sub)-1), TRUE),
                ref_line = 1,
                sizes = ifelse(logarithmic_scale, 0.6, 0.3),
                ci_column = 7,
                x_trans = ifelse(logarithmic_scale, "log", "none"),
                # xlim = ifelse(logarithmic_scale, xlim, NULL),
                # ticks_at = ifelse(logarithmic_scale, ticks_at, NULL),
                theme = raw_theme,
                footnote = paste0(df_text, "τ²=", tau_squared, ", I²=", i_squared, "%, p", p)
            )
            raw_forest_plot <- forestploter::edit_plot(raw_forest_plot, row = nrow(df_sub), gp = grid::gpar(fontface = "bold"))
            ggsave(paste0(folder_path, 'output/svg/raw-forest-plots--', output_file_name, '--', variable_group_value, '.svg'), raw_forest_plot, dpi = 300, width = 11.5)
            print('-----')
        }
        # sort by effect size (descending)
        new_rows <- new_rows[order(pooled_estimate, decreasing = TRUE)]
        if (significant_variables > 0) {
            if (!at_least_one_meta_analysis_performed_in_category & !is.na(subgroup_analysis)) {
                next
            }
            meta_analysis_output <- rbind(
                meta_analysis_output,
                data.table(
                    variable_group =  variable_category_value,
                    n_studies = "",
                    pooled_estimate = NA,
                    pooled_lower_95ci = NA,
                    pooled_upper_95ci = NA,
                    pooled_formatted_results = "",
                    se = NA
                ),
                use.names = FALSE
            )
            # add to the output table
            meta_analysis_output <- rbind(
                meta_analysis_output,
                new_rows,
                use.names = FALSE
            )
        }
    }
    if (at_least_one_meta_analysis_performed == F) {
        print('No meta-analysis performed')
        quit()
    }
    # make empty column for the actual forest plot
    meta_analysis_output$empty_column <- paste(rep(" ", 16), collapse = " ")
    # set column names and prepare data.table
    colname_measure_of_effect <- ifelse(logarithmic_scale, ifelse(ever_RR, ifelse(!only_RR, "OR or RR", "RR"), "OR"), "MD")
    # colname_measure_of_effect <- ifelse(logarithmic_scale, ifelse(OR_to_RR, "OR or RR", "OR"), "MD")
    if (analysis_type == 'risk-factors') {
        colnames(meta_analysis_output) <- c("Risk factors                       ", "k", "OR", "Lower 95%CI", "Upper 95%CI", paste0(colname_measure_of_effect, " (95%CI)"), "SE", " ")
    } else {
        colnames(meta_analysis_output) <- c("Outcomes                           ", "k", "OR", "Lower 95%CI", "Upper 95%CI", paste0(colname_measure_of_effect, " (95%CI)"), "SE", " ")
    }

    if (logarithmic_scale) {
        # set xlim and ticks_at
        x_lower_lim <- max(0, min(meta_analysis_output$`Lower 95%CI`, na.rm=T) - 0.1)
        x_lower_lim <- ifelse(x_lower_lim > 1, 0.9, x_lower_lim)
        xlim <- c(
            x_lower_lim,
            min(max(meta_analysis_output$`OR`, na.rm=T) + 2, max(meta_analysis_output$`Upper 95%CI`, na.rm=T))
        )
        geom_mean <- sqrt(min(xlim) * max(xlim))
        p10_val <- quantile(xlim, 0.3)
        geom_p10 <- sqrt(min(xlim) * p10_val)
        p90_val <- quantile(xlim, 0.7)
        geom_p90 <- sqrt(max(xlim) * p90_val)
        ticks_at <- c(geom_p10, geom_mean, geom_p90)
        ticks_at <- round(ticks_at, 1)
        print('ticks')
        print(ticks_at)
        print('meta-analysis')
        print(meta_analysis)
        # in case thre are any ticks within ±0.2 of 1, remove them and add one directly on 1, but if there are ≥2 such, don't do anything
        if (sum(abs(ticks_at - 1) <= 0.2) >= 1) {
            ticks_at <- ticks_at[abs(ticks_at - 1) >= 0.2]
            ticks_at <- c(ticks_at, 1)
        } else {
            ticks_at <- c(ticks_at, 1)
        }
        # if x_lim[1] is higher than 1, or x_lim[2] is lower than 1, set corresponding x_lim value to 1
        if (xlim[1] > 1) {
            xlim[1] <- 0.98
        }
        if (xlim[2] < 1) {
            xlim[2] <- 1.02
        }
        print('ticks_at')
        print(ticks_at)
        print('xlim')
        print(xlim)
    }

    # set up the theme for the forest plot
    forest_plot_theme <- forestploter::forest_theme(
        core = list(bg_params=list(fill = c("transparent"))),
        xaxis_lwd = 1.3,
        refline_lwd = 1.3,
        refline_lty = 1,
        base_size = 9
    )
    # generate forest plot
    forest_plot <- forestploter::forest(
            meta_analysis_output[, c(1,2,6,8)],
            est = meta_analysis_output$`OR`,
            lower = meta_analysis_output$`Lower 95%CI`,
            upper = meta_analysis_output$`Upper 95%CI`,
            # sizes = meta_analysis_output$se,
            sizes = ifelse(logarithmic_scale, 0.6, 0.3),
            ci_column = 4,
            ref_line = 1,
            # arrow_lab = c("text to left", "text to right"),
            # footnote = "footnote text (prob fill with heterogeneity measures etc)"),
            x_trans = ifelse(logarithmic_scale, "log", "none"),
            # xlim = ifelse(logarithmic_scale, xlim, NULL),
            # ticks_at = ifelse(logarithmic_scale, ticks_at, NULL),
            theme = forest_plot_theme
    )
    # save the forest plot as svg 300 dpi with ggsave
    ggsave(paste0(folder_path, 'output/svg/forest-plots-', output_file_name, '.svg'), forest_plot, dpi = 300, width = 9)
    cat('-------- end\n')

    # FDR correction
    print('FDR CORRECTION')
    p_values$p_value_raw <- as.numeric(p_values$p_value_raw)
    p_values$p_value_fdr <- format.pval(FDRestimation::p.fdr(p_values$p_value_raw)$fdrs, digits = 4)
    print(p_values)
    # save p_values to .xlsx file in output/xlsx
    openxlsx::write.xlsx(p_values, paste0(folder_path, 'output/xlsx/p-values--', output_file_name, '.xlsx'), rowNames = FALSE)
}