# Load packages
packages_all = c("readr", "ggplot2", "data.table", "stringr")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
subfolder <- 'AD-eczema' # change this to the desired subfolder
plot_name <- paste0(subfolder, '--', 'late-onset-late-resolving--probability') # change to this, this is the regular one
# plot_name <- paste0('wheezing', '--', 'early-onset-mid-resolving--prevalence') # use this for astma/wheezing
source(paste0(folder_path, 'input/r/', 'CLASSES-', plot_name, '.R'))
max_age <- 18
x_ticks_interval <- ifelse(max_age < 18, 2, 3)
line_width <- 1.3 # 1.77
show_legends <- F
height <- ifelse(show_legends, 5.5, 4.7)
# check if plot_name contains '--probability'
outcome <- ifelse(stringr::str_detect(plot_name, '--probability'), 'probability', 'prevalence')



# make a data.table to hold all plot data
plot_data_table <- data.table(
    trajectory = character(),
    x = numeric(),
    y = numeric()
)
for (plot_class in plot_classes) {
    # read the CSV file
    df <- readr::read_csv(paste0(folder_path, "input/csv/", subfolder, "/", plot_class, ".csv"), col_names = F, show_col_types = F, name_repair = "unique_quiet")
    # convert to data.table
    df <- as.data.table(df)
    # set the column names
    colnames(df) <- c("x", "y")
    # add the class name
    df$trajectory <- plot_class
    # add to the plot data table
    plot_data_table <- rbind(plot_data_table, df)
}
# # print plot_data_table by class
# for (plot_class in plot_classes) {
#     print(plot_data_table[trajectory == plot_class])
# }
print(plot_data_table)



# multi-disease
# for multimorbidity trajectories, to differentiate the same variable for different studies
if (plot_name == 'multiple-diseases--rhinitis-cough') {
    plot_line_colors <- c('#2b2b2b', '#932f1e', '#2b7135', '#2f4a8f', '#2b2b2b', '#932f1e', '#2b7135', '#2f4a8f')
    plot_line_types <- c(rep('solid', 4), rep('dashed', 4))
}
# atopy
# for early persistent milk and egg sensitization
if (plot_name == 'atopy--early-milk-egg') {
    plot_line_colors <- c('#f58231', '#e6194c', '#f58231', '#e6194c', '#f58231', '#e6194c')
    plot_line_types <- c(rep('solid', 2), rep('dashed', 2), rep('dotted', 2))
}
# for late mite sensitization
if (plot_name == 'atopy--late-mite') {
    plot_line_colors <- c('#0034b7', '#f58231', '#34723c', '#34723c')
    plot_line_types <- c(rep('solid', 3), 'dashed')
}
# AD/eczema
# for early-onset persistent AD-eczema (prevalence)
if (plot_name == 'AD-eczema--early-onset-persistent--prevalence') {
    plot_line_colors <- c('#7ecf37', '#7ecf37', '#7ecf37', '#8f2a8a', '#8f2a8a')
    plot_line_types <- c('solid', 'dashed', 'dotted', 'solid', 'dashed')
}
# for early-onset persistent AD-eczema (probability)
if (plot_name == 'AD-eczema--early-onset-persistent--probability') {
    plot_line_colors <- c('#0034b7', '#f58231', '#34723c', '#f032e6', '#aaaaaa', '#e6194c', '#9a6334', '#9a6334')
    plot_line_types <- c(rep('solid', 7), 'dashed')
}
# for early-onset early-resolving AD-eczema (prevalence)
if (plot_name == 'AD-eczema--early-onset-early-resolving--prevalence') {
    plot_line_colors <- c('#8f2a8a', '#8f2a8a', '#8f2a8a', '#e6194c', '#e6194c')
    plot_line_types <- c('solid', 'dashed', 'dotted', 'solid', 'dashed')
}
# asthma/wheezing
# for early-onset persisten wheeze (probability)
if (plot_name == 'wheezing--early-onset-persistent--probability') {
    plot_line_colors <- c('#585858', '#585858', '#f99d9e', '#000000', '#7cb848', '#4dc0bc', '#e6194c', '#800000', '#1978bd', '#8f2a8a', '#714516', '#4527b2')
    plot_line_types <- c('solid', 'dashed', rep('solid', 10))
}
# for early-onset persistent wheezing (prevalence)
if (plot_name == 'wheezing--early-onset-persistent--prevalence') {
    plot_line_colors <- c('#0034b7', '#f58231', '#34723c', '#34723c', '#f032e6', '#aaaaaa')
    plot_line_types <- c(rep('solid', 3), 'dashed', 'solid', 'solid')
}
# for mid-onset wheezing (probability)
if (plot_name == 'wheezing--mid-onset--probability') {
    plot_line_colors <- c('#0034b7', '#f58231', '#34723c', '#f032e6', '#aaaaaa', '#aaaaaa', '#9a6334', '#9a6334', '#414141')
    plot_line_types <- c(rep('solid', 5), 'dashed', 'solid', 'dashed', 'solid')
}
# for mid-onset wheezing (prevalence)
if (plot_name == 'wheezing--mid-onset--prevalence') {
    plot_line_colors <- c('#60bedb', '#f99d9e', '#000000', '#78bc3d', '#78bc3d', '#78bc3d', '#800000')
    plot_line_types <- c(rep('solid', 4), 'dashed', 'dotted', 'solid')
}
# for early-onset mid-resolving wheezing (probability)
if (plot_name == 'wheezing--early-onset-mid-resolving--probability') {
    plot_line_colors <- c('#0034b7', '#f58231', '#c41717', '#34723c', '#f032e6', '#f032e6', '#aaaaaa', '#9a6334', '#f99d9e', '#414141')
    plot_line_types <- c(rep('solid', 5), 'dashed', rep('solid', 4))
}
# for early-onset mid-resolving wheezing (prevalence)
if (plot_name == 'wheezing--early-onset-mid-resolving--prevalence') {
    plot_line_colors <- c('#60bedb', '#60bedb', '#000000', '#901540', '#4527b2', '#4527b2', '#4527b2', '#8f2a8a', '#6e553b')
    plot_line_types <- c('solid', 'dashed', rep('solid', 3), 'dashed', 'dotted', 'solid', 'solid')
}
# for early-onset persistent wheezing (prevalence and probability; non-repeated colors) TODO!!!



# plot the data with one line per class ($trajectory)
if (outcome == 'probability') {
    plot_data_table$y <- plot_data_table$y / 100
}
plot <- ggplot(plot_data_table, aes(x = x, y = y, group = trajectory, color = trajectory, linetype = trajectory)) +
    geom_line(linewidth = line_width) +
    geom_point(size = 3) +
    labs(
        title = plot_name,
        x = "Age (years)",
        y = ifelse(outcome == 'probability', "Probability of outcome", "Prevalence of outcome (%)")
    ) +
    theme(
        panel.background = element_rect(fill = "transparent"),
        axis.line = element_line(color='black', linewidth = 0.75),
        axis.text.x = element_text(size = 16, color = "black", margin = margin(t = 5)),
        axis.text.y = element_text(size = 16, color = "black", margin = margin(r = 5)),
        axis.title.x = element_text(size = 18, color = "black", margin = margin(t = 9)),
        axis.title.y = element_text(size = 18, color = "black", margin = margin(r = 9)),
        axis.ticks = element_line(linewidth=0.65),
        axis.ticks.length = unit(0.2, "cm"),
        legend.position = ifelse(show_legends, "bottom", "none"),
        legend.title = element_blank(),
        legend.text = element_text(size  = 13),
        legend.margin = margin(10, 0, 10, 0),
        legend.key = element_blank(),
        legend.direction = "horizontal",
        # plot.title = element_text(size = 18, hjust = 0.5, margin = margin(b = 12)),
        plot.title = element_blank(),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank()
    ) +
    guides(color = guide_legend(ncol = 3)) +
    scale_x_continuous(breaks = seq(from = 0, to = max_age, by = x_ticks_interval), limits = c(0, max_age + .4), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = ifelse(outcome == 'probability', 1, 100), by = ifelse(outcome == 'probability', 0.3, 33)), limits = c(0, ifelse(outcome == 'probability', 1.05, 105)), expand = c(0, 0)) +
    scale_color_manual(values = plot_line_colors) +
    scale_linetype_manual(values = plot_line_types)
ggsave(
    paste0(folder_path, "output/svg/", "line-plot-", plot_name, ".svg"),
    plot,
    width = 10,
    height = height,
    dpi = 300
)