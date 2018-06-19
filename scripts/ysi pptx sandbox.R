

devtools::install_github("davidgohel/officer")



library('tidyverse')
library('officer')


# read from starter Template (has the logo in top corner)
p  <- read_pptx("docs/ysi.pptx")

p  <- p %>%

  # the title slide
  on_slide(index = 1) %>%
  ph_with_text(type = "ctrTitle", str = "YSI Presentation") %>%
  ph_with_text(type = "subTitle", str = "A Template") %>%

  # add three blank slides
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  add_slide(layout = "Title and Content", master = "Office Theme") %>%

  # on slide 2 add some text
  on_slide(index = 2) %>%
  ph_with_text(type = "title", str = "First title") %>%
  ph_with_text(type = "body", str = "Some text goes here")

# create some plots
pl1 <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length, col = Species), size = 3)

pl2 <- ggplot(data = mtcars) +
    geom_point(mapping = aes(mpg, disp, col = as.factor(gear)), size = 3)

# add them to slides 3 and 4
p <- p  %>%
  on_slide(index = 3) %>%
  ph_with_text(type = "title", str = "Third title") %>%
  ph_with_gg(value = pl1) %>%
  on_slide(index = 4) %>%
  ph_with_text(type = "title", str = "Fourth title") %>%
  ph_with_gg(value = pl2)


# use the print function to write to file
print(p, target = "docs/test.pptx")
