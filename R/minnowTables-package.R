#' @keywords internal
"_PACKAGE"

## usethis namespace: start
## usethis namespace: end
NULL

# columns referenced via dplyr NSE, and the spec lists that
# record_legend_shapes()/record_note_runs() append to in the global environment
utils::globalVariables(c("id", "end", "start", "legend_shape_specs", "note_run_specs"))
