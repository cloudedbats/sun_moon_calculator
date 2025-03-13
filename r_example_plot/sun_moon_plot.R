require(tidyverse)
require(readxl)
library(hms)
library(janitor)

# Clear workspace.
rm(list=ls())

# Import -----------------------------------------------------------------------

sun_moon = read_excel("data_in/sun_moon_ottenby_2025.xlsx")

# Transform --------------------------------------------------------------------

sun_moon = sun_moon |>
  janitor::clean_names() |>
  rename(local_date = date)
  
# Visualise --------------------------------------------------------------------

# Center y data at midnight instead of noon. 
y_limits = c(as_datetime(hm("15:00")),as_datetime(ymd_hm("1970-01-02 9:00")))

# Same for data.
sunset_sunrise = sun_moon |>
  select(local_date, sunset, dusk, dawn, sunrise) |>
  mutate(sunset = as_hms(sunset),
         dusk = as_hms(dusk),
         sunrise = as_datetime(as_datetime("1970-01-02") + as_hms(sunrise)),
         dawn = as_datetime(as_datetime("1970-01-02") + as_hms(dawn)),
         local_date = as_date(local_date)
  )

plot_sunset_sunrise <- ggplot() +
  geom_line(data = sunset_sunrise,
            aes(x = local_date, y = sunset), size=0.1, colour="black") +
  geom_line(data = sunset_sunrise,
            aes(x = local_date, y = sunrise), size=0.1, colour="black") +
  geom_line(data = sunset_sunrise,
            aes(x = local_date, y = dusk), size=0.1, colour="grey") +
  geom_line(data = sunset_sunrise,
            aes(x = local_date, y = dawn), size=0.1, colour="grey") +
  scale_x_date(date_labels = "%Y-%m-%d") +
  scale_y_time(
    labels = function(t) strftime(t, '%H:%M'),
    limits = y_limits) +
  labs(x = "Monitoring nights", y = "Time for detected bats") +
  theme_bw()

plot_sunset_sunrise

# Export -----------------------------------------------------------------------

dir_results = "results"
# Create folder if not exists.
if (!dir.exists(dir_results)) {
  dir.create(dir_results)
}

plot_sunset_sunrise
ggsave(
  file.path(dir_results, "sun_ottenby_2025_plot.png"),
)

