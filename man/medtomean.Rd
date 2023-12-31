% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/data.functions.R
\name{medtomean}
\alias{medtomean}
\title{Converts medians to means using Box-Cox transformation}
\usage{
medtomean(df, scale = NULL, measurement = "base")
}
\arguments{
\item{df}{A dataframe in which continuous variables are labelled with the following
information (each separated by an underscore (_)):
\itemize{
\item scale: the name of the measurement scale (e.g. "beck") (if \code{scale=NULL} then this isn't necessary)
\item measurement: the time-point of interest (e.g. "cfb", "base", "final")
\item moment: the moment required for data transformation. Moments required in \code{df} for
transformation are \code{median}, \code{iqrlow}, \code{iqrhigh}.
}}
}
\description{
Converts medians to means using Box-Cox transformation
}
\details{
An example of the variables included in \code{df} may be \code{"beck_cfb_median"}, or
\code{"hamd_base_iqrlow"}. Means and SDs estimated using this function will then be stored as new
variables (or will overwrite existing ones) corresponding to \code{scale_measurement_mean} and \code{scale_measurement_sd}
}
