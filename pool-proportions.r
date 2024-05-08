# Load packages
packages_all = c("data.table", "meta", "readxl", "forestploter", "ggplot2")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
input_files <- c(
    'meta-analysis--wheezing--early-onset-mid-resolving--characteristics'
)



for (input_file in input_files) {
    cat('NEW FILE\n')
    cat('-------- start\n')
    # Read the Excel file
    df <- readxl::read_excel(paste0(folder_path, "output/docx/", input_file, ".xlsx"))
    df <- as.data.frame(df)
    df$variable <- sub("^(.)", "\\U\\1", df$variable, perl = TRUE)
    df$variable_category <- sub("^(.)", "\\U\\1", df$variable_category, perl = TRUE)
    print(paste('Loaded data from', input_file))

    # make sure to loop categories in the order they are defined
    predefined_ordered_variable_categories <- c(
        'Sociodemographics',
        'Prenatal exposure / hereditary',
        'Early-life exposure',
        'Other'
    )
    present_ordered_variable_categories <- c()
    for (variable_category_value in predefined_ordered_variable_categories) {
        if (variable_category_value %in% unique(df$variable_category)) {
            present_ordered_variable_categories <- c(present_ordered_variable_categories, variable_category_value)
        }
    }
    # data.table to hold textual summary of the meta-analysis
    summary_text <- c()
    # loop through the categories
    for (present_variable_category in present_ordered_variable_categories) {
        summary_text <- c(summary_text, present_variable_category)
        cat(present_variable_category, '\n')
        unique_variables <- unique(df$variable[df$variable_category == present_variable_category])
        unique_variables <- unique_variables[order(unique_variables)]
        # for each variable (column: variable), run meta-analysis
        for (unique_variable in unique_variables) {
            # set df_sub to get the subset of df in which variable is unique_variable
            df_sub <- df[df$variable == unique_variable,]
            df_sub$n_positive <- as.numeric(df_sub$n_positive)
            df_sub$n_class <- as.numeric(df_sub$n_class)
            ma <- meta::metaprop(
                df_sub$n_positive,
                df_sub$n_class,
                df_sub$data_id,
                data = df_sub,
                sm = "PLOGIT",
                method = "GLMM",
                method.ci ="WS",
                comb.fixed = FALSE
            )
            meta::forest(ma)
            # process summary data
            tau_squared <- ma$tau2
            i_squared <- ma$I2
            p_value <- ma$pval.Q[1]
            if (p_value < 0.01) {
                p_value <- '<0.01'
            } else {
                p_value <- paste0('=', round(p_value, 2))
            }
            footnote_text <- paste0("τ²= ", round(tau_squared, 1), ", I²= ", round(i_squared, 1), ", p", p_value)
            # backtransform pooled data
            random_estimates <- c(ma$TE.random,ma$lower.random,ma$upper.random)
            backtransformed_ma_data <- meta:::backtransf(random_estimates, sm="PLOGIT")
            point_estimate <- round(backtransformed_ma_data[1]*100,1)
            lower_95ci <- round(backtransformed_ma_data[2]*100,1)
            upper_95ci <- round(backtransformed_ma_data[3]*100,1)
            formatted_pooled_data <- paste0(point_estimate, " (", lower_95ci, ", ", upper_95ci, ")")
            # addition_summary_text <- paste0(unique_variable, ": ", formatted_pooled_data, " ", paste0(df_sub$data_id, collapse = ", ")) # for extra details
            addition_summary_text <- paste0(unique_variable, ": ", formatted_pooled_data)
            cat(addition_summary_text, '\n')
            summary_text <- c(
                summary_text,
                addition_summary_text
            )
            individual_point_estimates <- round(meta::backtransf(ma$TE, sm="PLOGIT")*100,1)
            individual_lower_95ci <- round(meta::backtransf(ma$lower, sm="PLOGIT")*100,1)
            individual_upper_95ci <- round(meta::backtransf(ma$upper, sm="PLOGIT")*100,1)
            individual_formatted_data <- paste0(individual_point_estimates, " (", individual_lower_95ci, ", ", individual_upper_95ci, ")")
            forest_plot_data <- data.table()
            forest_plot_data <- data.table(
                data_id = df_sub$data_id,
                point_estimate = individual_point_estimates,
                lower_95ci = individual_lower_95ci,
                upper_95ci = individual_upper_95ci,
                n_class = df_sub$n_class,
                formatted_data = individual_formatted_data,
                empty_colum = paste(rep(" ", 33), collapse = " ")
            )
            # add summary data to the forest plot data
            forest_plot_data <- rbind(
                forest_plot_data,
                data.table(
                    data_id = 'Pooled',
                    point_estimate = point_estimate,
                    lower_95ci = lower_95ci,
                    upper_95ci = upper_95ci,
                    n_class = sum(df_sub$n_class),
                    formatted_data = formatted_pooled_data,
                    empty_colum = paste(rep(" ", 33), collapse = " ")
                )
            )
            xlim <- c(
                0,
                min(100, max(forest_plot_data$upper_95ci, na.rm = TRUE) + 5)
            )
            # ticks_at should be three in quantity, one in the middle and two at the ends (33% from xlim start and 66% from xlim end)
            ticks_at <- c(
                xlim[1] + (xlim[2] - xlim[1]) * 0.15,
                xlim[1] + (xlim[2] - xlim[1]) * 0.5,
                xlim[1] + (xlim[2] - xlim[1]) * 0.85
            )
            ticks_at <- round(ticks_at, 1)
            # make forest plot
            forest_theme <- forestploter::forest_theme(
                core = list(bg_params=list(fill = c("transparent"))),
                xaxis_lwd = 1.3,
                refline_lwd = 1.3,
                refline_lty = 2,
                base_size = 9,
                footnote_cex = 0.75,
                summary_fill = "#575e92",
                summary_col = "#575e92"
            )
            colnames(forest_plot_data) <- c("Data ID", "Proportion", "Lower 95%CI", "Upper 95%CI", "N(class)", "Proportion (%) (95%CI)", " ")
            forest_plot <- forestploter::forest(
                forest_plot_data[, c(1, 5:7)],
                est = forest_plot_data$`Proportion`,
                lower = forest_plot_data$`Lower 95%CI`,
                upper = forest_plot_data$`Upper 95%CI`,
                is_summary = c(rep(FALSE, nrow(forest_plot_data)-1), TRUE),
                ref_line = point_estimate,
                sizes = 0.72,
                ci_column = 4,
                xlim = xlim,
                ticks_at = ticks_at,
                theme = forest_theme,
                footnote = footnote_text
            )
            forest_plot <- forestploter::edit_plot(forest_plot, row = nrow(forest_plot_data), gp = grid::gpar(fontface = "bold"))
            suppressMessages(ggsave(paste0(folder_path, 'output/svg/raw-forest-plots--', input_file, '--', unique_variable, '.svg'), forest_plot, dpi = 300, width = 10))
            print('-----')
        }
    }
    # save summary_text to a .txt file
    write(summary_text, file = paste0(folder_path, 'output/txt/trajectory-characteristics-summary--', input_file, '.txt'))
    cat('-------- end\n')
}