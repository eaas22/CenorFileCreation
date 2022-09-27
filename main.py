from Apps.file_manager import file_creation
from gooey import Gooey, GooeyParser


@Gooey(required_cols=1,
       program_name="CENOR File Creation",
       program_description="Script to run an excel file creation",
       clear_before_run=True, )
def parse_args():
    desc = "Philips DA"
    cockpit_help_msg = "Select the cockpit file"
    master_data_help_msg = "Select the Logistic Master Data file"
    number_help_msg = "Enter here the expedition number"

    # Constructing the GUI
    my_parser = GooeyParser(description=desc)
    my_parser.add_argument(
        "cockpit_input", action="store", help=cockpit_help_msg, widget="FileChooser",
        gooey_options=dict(wildcard="MS Excel Files|*.xlsx"))
    my_parser.add_argument(
        "logi_md", action="store", help=master_data_help_msg, widget="FileChooser",
        gooey_options=dict(wildcard="MS Excel Files|*.xlsx"))
    my_parser.add_argument(
        "expedition_number", help=number_help_msg, widget="TextField")

    arguments = my_parser.parse_args()
    return arguments


if __name__ == '__main__':
    args = parse_args()
    cockpit_input = args.cockpit_input
    logi_md = args.logi_md
    expedition_number = args.expedition_number
# try:
    file_creation(cockpit_input, logi_md, expedition_number)
# except Exception as e:
#     print(e)
