import openpyxl
import plotly.graph_objects
import openpyxl.utils
import numbers
from state_abbrev import us_state_to_abbrev


#this is the main function where I use all my functions together
def main():
    pop_worksheet = open_worksheet()
    show_pop_change = should_display_pop_change()
    if show_pop_change:
        show_pop_change_map(pop_worksheet)
    else:
        show_percent_change_map(pop_worksheet)


# this function basically opens up the Excel document and returns its data
def open_worksheet():
    income_excel = openpyxl.load_workbook("countyPopChange20202021.xlsx")
    data_sheet = income_excel.active
    return data_sheet


# ask the user if the program should display a map of total population changes
# and returns a true or false value depending on the user's response
def should_display_pop_change():
    show_display = input("Would you like to display a map of total population changes?")
    no = ["no", "nah", "nope"]  # this is a list of possible answers the user could give back that are equivalent to no
    yes = ["yes", "ya",
           "yeah"]  # this is a list of possible answers the user could give back that are equivalent to yes
    if show_display in no:
        return False
    if show_display in yes:
        return True


# for the yes responses:
def show_pop_change_map(population_sheet):
    list_of_state_abrev = []
    list_of_npopchg2021 = []
    for row in population_sheet.rows:
        first_cell = row[5]
        population_cell = row[11]
        population_value = population_cell.value
        state_name = first_cell.value
        if not isinstance(population_value, numbers.Number):
            continue
        if state_name not in us_state_to_abbrev:
            continue
        state_abrev = us_state_to_abbrev[state_name]
        list_of_state_abrev.append(state_abrev)
        npopchg2021_cell_number = openpyxl.utils.cell.column_index_from_string('l') - 1
        npopchg2021_cell = row[npopchg2021_cell_number]
        npopchg2021 = npopchg2021_cell.value
        list_of_npopchg2021.append(npopchg2021)

    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations=list_of_state_abrev,
            z=list_of_npopchg2021,
            locationmode="USA-states",
            colorscale='Portland',
            colorbar_title="Total Population Change"
        )
    )
    map_to_show.update_layout(
        title_text="Total Population Change in 2021",
        geo_scope="usa"
    )
    map_to_show.show()


# for the no responses:
def show_percent_change_map(population_sheet):
    list_of_state_abrev = []
    list_of_pop_change_percent = []
    for row in population_sheet.rows:
        first_cell = row[5]
        population_change_cell = row[11]
        pope_estimate_cell = row[9]
        population_change_value = population_change_cell.value
        pope_estimate_value = pope_estimate_cell.value
        state_name = first_cell.value
        if not isinstance(population_change_value, numbers.Number):
            continue
        if not isinstance(pope_estimate_value, numbers.Number):
            continue
        if state_name not in us_state_to_abbrev:
            continue
        state_abrev = us_state_to_abbrev[state_name]
        list_of_state_abrev.append(state_abrev)
        npopchg2021_cell_number = openpyxl.utils.cell.column_index_from_string('l') - 1
        popestimate2021_cell_number = openpyxl.utils.cell.column_index_from_string('j') - 1
        npopchg2021_cell = row[npopchg2021_cell_number]
        popestimate2021_cell = row[popestimate2021_cell_number]
        npopchg2021 = npopchg2021_cell.value
        popestimate2021 = popestimate2021_cell.value
        pop_change_percent = npopchg2021 / popestimate2021
        pop_change_percent = pop_change_percent * 100
        list_of_pop_change_percent.append(pop_change_percent)

    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations=list_of_state_abrev,
            z=list_of_pop_change_percent,
            locationmode="USA-states",
            colorscale='Picnic',
            colorbar_title="Percentage of Population Change"
        )
    )
    map_to_show.update_layout(
        title_text="Percentage Change of Population from Total Population to the 2021 Population Change",
        geo_scope="usa"
    )
    map_to_show.show()


main()