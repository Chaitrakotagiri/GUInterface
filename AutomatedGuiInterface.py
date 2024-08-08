import os.path
import PySimpleGUI as sg
from datetime import datetime
import threading
import json
from csv import writer
import matplotlib.pyplot as plt
import rs_instr
import logging
import tkinter as mytk
import pygetwindow as gw
from PIL import ImageGrab
import pandas as pd
import openpyxl
import yaml
import time
import sys


SW_VERSION = '0.0'
timestamp = time.strftime("%Y%m%d_%H%M%S")

combinations_file_path = os.getcwd()
combinations_file_name = "combinations.json"
saved_combinations = []

try:
    with open(os.path.join(combinations_file_path, combinations_file_name), "r") as combinations_file:
        file_contents = combinations_file.read()
        if file_contents:
            saved_combinations = json.loads(file_contents)
        else:
            print("The file is empty.")
except FileNotFoundError:
    saved_combinations = []
except json.JSONDecodeError as e:
    print(f"Error decoding JSON: {e}")

MOP = 24
test_data = {
    'Tests': {
        'Test1': [
            {
                'name': 'test_1',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'Carrier': [1, 2, 3, 4, 5, 6],
                'lf': True,
                'ui_power': False,
                'mod': False,
                'ca': False,
                'carrier': True,
                'op': False
            },
        ],

        'Test2': [
            {
                'name': 'test_2',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'Carrier': [1, 2, 3, 4, 5, 6],
                'lf': True,
                'ui_power': True,
                'mod': False,
                'ca': False,
                'carrier': True,
                'op': False
            },
        ],
        'Test3': [
            {
                'name': 'test_3',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'modulations': ['QPSK', '8PSK', '16QAM'],
                'CA': ['Single', 'Dual'],
                'Carrier': [1, 2, 3, 4, 5, 6],
                'lf': True,
                'ui_power': True,
                'mod': True,
                'ca': True,
                'carrier': True,
                'op': False
            },
        ],
        'Test4': [
            {
                'name': 'test_4',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'modulations': ['QPSK', '8PSK', '16QAM'],
                'CA': ['Single', 'Dual'],
                'Carrier': [1, 2, 3, 4, 5, 6],
                'lf': True,
                'ui_power': True,
                'mod': True,
                'ca': True,
                'carrier': True,
                'op': False
            },
        ],
        'Test5': [
            {
                'name': 'test_5',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'Carrier': [1, 2, 3, 4, 5, 6, 'NA'],
                'lf': True,
                'ui_power': True,
                'mod': False,
                'ca': False,
                'carrier': True,
                'op': True
            },
        ],

        'Test6': [
            {
                'name': 'test_6',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'Carrier': [1, 2, 3, 4, 5, 6, 'NA'],
                'lf': True,
                'ui_power': True,
                'mod': False,
                'ca': False,
                'carrier': True,
                'op': True
            },
        ],
        'Test7': [
            {
                'name': 'test_7',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'power_levels': [MOP-2, MOP-1, MOP, MOP+1, MOP+2],
                'modulations': ['CW'],
                'Carrier': ['NA'],
                'lf': True,
                'ui_power': True,
                'mod': True,
                'ca': False,
                'carrier': True,
                'op': False
            },
        ],

        'Test8': [
            {
                'name': 'test_8',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'lf': True,
                'ui_power': False,
                'mod': False,
                'ca': False,
                'carrier': False,
                'op': False
            },
                ],

        'Test9': [
            {
                'name': 'test_9',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'lf': True,
                'ui_power': False,
                'mod': False,
                'ca': False,
                'carrier': False,
                'op': False
            },
        ],

        'Test10': [
            {
                'name': 'test_10',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'lf': True,
                'ui_power': False,
                'mod': False,
                'ca': False,
                'carrier': False,
                'op': False
            },
        ],

        'Test11': [
            {
                'name': 'test_11',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'modulations': ['QPSK', '16QAM', 'CW'],
                'lf': True,
                'ui_power': False,
                'mod': True,
                'ca': False,
                'carrier': False,
                'op': False
            },
        ],

        'Test12': [
            {
                'name': 'test_12',
                'temperatures': [-45, 25, 50],
                'LOFreq': [10000, 10050, 10100, 10150],
                'lf': True,
                'ui_power': False,
                'mod': False,
                'ca': False,
                'carrier': False,
                'op': False
            },
        ],

    }
}

class AutomatedGui():
    def __init__(self):
        self.width, self.height = self.get_window_size()
        if self.height >= 1080:
            self.height = 950
        if self.width >= 1920:
            self.width = 1200

        print(f"Screen res is {self.width}x{self.height}")
        self.selected_combinations = []
        self.saved_combinations = []
        self.running_combinations = False
        self.pause_combinations = False
        self.simu_mode = False
        self.selected_parameters_list = []
        self.test_data = test_data
        self.test_list = sg.DropDown([], size=(20, 8), key="-TESTLIST-")
        self.checkbox_to_param = {
            "temp": "temperatures",
            "lofreq": "LOFreq",
            "carrier": "Carrier",
            "power_level": "power_levels",
            "modulations": "modulations",
            "CA": "CA",
            "Option": "Option"
        }
        self.report_str = ""
        self.dut_name = ""
        self.dut_vendor = ""
        self.dut_rev = ""
        self.dut_sn = ""
        #self.tempr = 25
        self.dut_info_file = "dut_info.yaml"
        if self.dut_name == "" and self.dut_vendor =="" and self.dut_rev == "" and self.dut_sn == "":
            try:
                with open(self.dut_info_file, "r") as file:
                    data = yaml.safe_load(file)
                    self.dut_name = data.get("dut_name", "")
                    self.dut_vendor = data.get("dut_vendor", "")
                    self.dut_rev = data.get("dut_rev", "")
                    self.dut_sn = data.get("dut_sn", "")
                    self.operator = data.get("Operator", "")
            except FileNotFoundError as e:
                print(f"Could not find the file, {e}")
        self.fname_prefix = self.dut_name + "_" + self.dut_vendor + "_" + self.dut_rev + "_" + self.dut_sn

        time_stamp = datetime.now()
        date_time = time_stamp.strftime("%Y-%m-%d_%H-%M-%S")
        self.terminal = sys.stdout

        w = self.width
        h = self.height

        w_char = w // 9
        h_char = h // 50

        font = 'Calibri'
        font_size = 13
        layout = []
        tab_layout = []
        menu_bar = [['File', ['Save', 'Exit', 'DUT Configure']],
                    ['Help', ['About']],]
        calibration_options = ["Path 1", "Path 2", "Path 3", "Path 4"]
        status_frame_layout = [
            [
                sg.Button("DUT Configure", key='DUT Configure'),
            ],
            [sg.Text(f'DUT:{self.dut_name} DUT Type: {self.dut_vendor}',font=(font, 12, 'bold')), sg.Text(key='UnitType', size=(3, 1), font=(font, 10, 'bold')),
             sg.Text(f'DUT SN:{self.dut_sn}', font=(font, 10, 'bold')), sg.Text(key='dut_sn', size=(3, 1) , font=(font, 10, 'bold')),
            ]
        ]
        test_equipment_status_layout = [
            [
                sg.Button("Check Equipment Status", key='Check_equip_status'),
            ],

            [
                sg.Text(f'Signal Generator', font=(font, 12, 'bold')),
                sg.Text('SG_status'), self.LEDIndicator('SG_status'),

             sg.Text(f'Spectrum Analyser', font=(font, 12, 'bold')),
             sg.Text('SA_status'), self.LEDIndicator('SA_status'),
             ]


        ]
        tab_layout_first = []
        tab_layout_second = []
        index = 0
        for section_name, tests in test_data['Tests'].items():
            index += 1
            # tab_section = []
            first_tab_section = []
            second_tab_section  = []
            len_tests = len(test_data['Tests'])
            half_value = len_tests//2


            # Iterate through the tests in each section
            for test in tests:
                test_name = test['name']

                # Create a section for each test
                test_section = [
                    sg.Frame(f"Test Name: {section_name}",[

                        [sg.Text("Option", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(Option, tooltip='Right click for more test options' if i == 0 else '',
                                     right_click_menu=['Options',['Narrow Band', 'Wide Band', 'Both']] if i == 0 else [],
                                     default=False, enable_events=True, key=f"Option_{test_name}_{Option}", font=(font, font_size, 'bold')) for
                         i, Option in enumerate(test.get('Option', []))] if test.get('op', False) else [],

                        [sg.Text("Temperatures (C)", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(temp, key=f"temp_{test_name}_{temp}", font=(font, font_size, 'bold')) for temp in
                         test.get('temperatures', [])],

                        [sg.Text("LOFreq (MHz)", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(lofreq, key=f"lofreq_{test_name}_{lofreq}", font=(font, font_size, 'bold')) for lofreq in
                         test.get('LOFreq', [])]+
                        [sg.Text("User Input (MHz):", font=(font, font_size, 'bold'))] +
                        [sg.InputText("", size=(5, 1), key=f"lofreq_input_{test_name}")],

                        [sg.Text("Output Power Levels (dB)", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(pl, key=f"power_level_{test_name}_{pl}", font=(font, font_size, 'bold')) for pl in
                         test.get('power_levels', [])]+
                        [sg.Text("User Input (dB):",font=(font, font_size, 'bold'))]+
                        [sg.InputText("", size=(5, 1), key=f"power_input_{test_name}")] if test.get('ui_power', False) else [],

                        [sg.Text("Modulations", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(mod, key=f"modulations_{test_name}_{mod}", default=False, enable_events=True, font=(font, font_size, 'bold')) for mod in
                         test.get('modulations', [])] if test.get('mod', False) else [],

                        [sg.Text("CA", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(CA, default= False, enable_events=True, key=f"CA_{test_name}_{CA}", font=(font, font_size, 'bold')) for CA in
                         test.get('CA', [])] if test.get('ca', False) else [],

                        [sg.Text("Carrier", font=(font, font_size, 'bold'))] +
                        [sg.Checkbox(carrier, key=f"carrier_{test_name}_{carrier}", font=(font, font_size, 'bold')) for carrier in
                         test.get('Carrier', [])] if test.get('carrier', False) else [],

                    ], font=(font, 14, 'bold'), border_width=3, title_color='Yellow'),
                ]
                if index <= half_value:
                    first_tab_section.append(test_section)
                else:
                    second_tab_section.append(test_section)

            if first_tab_section:
                first_tab = sg.Tab(section_name, first_tab_section)
                tab_layout_first.append(first_tab)
            else:
                second_tab = sg.Tab(section_name, second_tab_section)
                tab_layout_second.append(second_tab)

            #     tab_section.append(test_section)
            # tab = sg.Tab(section_name, tab_section)
            # tab_layout.append(tab)

        test_configure_options = [[sg.Button("Load Earlier Tests", key="-LOAD_COMBINATIONS-"),
                       sg.Text("Calibration Option:", font=(font, 10, 'bold')),
                       sg.DropDown(calibration_options, size=(20, 8), key="-CALIBRATION_OPTION-"),
                       sg.Button("Calibrate Loss", key="-CALIBRATE-")], [sg.Button("Add to Test List"), sg.Button("View"),
                       sg.Button("View All Test Cases", key="-VIEW_ALL_TEST_CASES-"), sg.Button("Clear Selections"),
                       sg.Button("Delete All Test Cases"), sg.Button("Delete Test Case")], [self.test_list, sg.Button("Open Loss Calibration File")]]

        run_test_layout = [[sg.Button("Run Selected Test In Testlist", button_color='Green', size=(8,4)), sg.Button("Run All Tests"),
             sg.Button("Pause"), sg.Button("Resume"), sg.Button("Stop"), sg.Button("Exit")]]

        layout.append([sg.Menu(menu_bar)])
        layout.append([sg.Frame('Current DUT status', status_frame_layout, title_color='black', font=(font, 12, 'bold')), sg.Frame('Current Equipment status', test_equipment_status_layout, title_color='black', font=(font, 12, 'bold'))])
        layout.append([sg.Frame('Test Config Options',test_configure_options, title_color='black', font=(font, 12, 'bold')), sg.Frame('Test Run Window',run_test_layout, title_color='black', font=(font, 12, 'bold')) ])
        # layout.append([sg.TabGroup([tab_layout], key="-PARAMS-", focus_color='White', font=(font, font_size, 'normal'),
        #                            tab_location='left', tab_border_width=1, expand_x=True, expand_y=True,
        #                            selected_title_color='Black', selected_background_color='Yellow')])
        layout.append([sg.TabGroup([
            [sg.Tab('First Set Tests', [[sg.TabGroup([tab_layout_first], font=(font, font_size, 'normal'),
                                              selected_background_color='Yellow', expand_x=True, expand_y=True,
                                              selected_title_color='Black', tab_border_width=1, focus_color='White',
                                              tab_location='left', key="-PARAMS-", )]]),
             sg.Tab('Second Set Tests', [[sg.TabGroup([tab_layout_second], font=(font, font_size, 'normal'),
                                              selected_background_color='Yellow', expand_x=True, expand_y=True,
                                              selected_title_color='Black', tab_border_width=1, focus_color='White',
                                              tab_location='left', key="-PARAMS-", )]],
                    font=(font, font_size, 'normal'))],
        ], font=(font, 14, 'normal'), selected_title_color='Black', selected_background_color='Yellow')])
        layout.append([sg.Multiline("", key="-OUTPUT-", size=(110, 14))])
        layout.append([sg.Button("Clear Output", key="-CLEAR-")])
        layout.append([
            sg.Text("Sub Progress", font=("", 15, "bold")),
            sg.ProgressBar(1000, orientation='h', size=(w_char - 75, 20), key='sub-progress',
                           bar_color=(sg.YELLOWS[0], 'grey'), expand_x=True)]
        )
        layout.append([
                                  sg.Text("Progress", font=("", 15, "bold")),
                                  sg.ProgressBar(1000, orientation='h', size=(w_char - 75, 20), key='progress',
                                             bar_color=(sg.YELLOWS[0], 'grey'), expand_x=True)]
                      )
        self.window = sg.Window(f'Automated Test GUI - SW Version: {SW_VERSION} ', layout, finalize=True, size=(self.width, self.height), resizable=True)
        self.original_size = self.window.size

    def LEDIndicator(self, key=None, radius=30):
        return sg.Graph(canvas_size=(radius, radius),
                        graph_bottom_left=(-radius, -radius),
                        graph_top_right=(radius, radius),
                        pad=(0, 0), key=key)

    def SetLED(self, window, key, color, text):
        graph = window[key]
        graph.erase()
        graph.draw_circle((0, 0), 30, fill_color=color, line_color=color)
        graph.draw_text(text, (0, 0), font = ("", 4, "bold"))

    def load_from_yaml(self):
        try:
            with open("dut_info.yaml", "r") as file:
                data = yaml.safe_load(file)
                self.dut_name = data.get("dut_name", "")
                self.dut_vendor = data.get("dut_vendor", "")
                self.dut_rev = data.get("dut_rev", "")
                self.dut_sn = data.get("dut_sn", "")
        except FileNotFoundError as e:
            print(f"Could not find the file, {e}")
        return self.dut_name, self.dut_vendor, self.dut_rev, self.dut_sn

    def write(self, message): # This is for the log
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):  # This is for the log
        self.terminal.flush()
        self.log.flush()

    def mainloop(self): # This loop handles all the events of the GUI
        w = self.width
        h = self.height
        combinations_thread = None
        captured_data = []

        # Event loop
        while True:
            selected_calibration_option = None
            (event, value) = self.window.Read(timeout=10)
            # print(f"Event: {event}, Values: {value}")

            if event == sg.WIN_CLOSED or event == "Exit":
                with open(os.path.join(combinations_file_path, combinations_file_name), "w") as combinations_file:
                    json.dump(self.selected_combinations, combinations_file)
                break

            if event == "-LOAD_COMBINATIONS-":
                self.load_combinations()

            # This event is for the checkboxes to hide in case if selected few options

            if event == "-CALIBRATE-":
                selected_calibration_option = value["-CALIBRATION_OPTION-"]
                calibration_result = None

                if selected_calibration_option == "Path 1":
                    print("Selection option Path 1")
                    # Placeholder for what has to be done
                elif selected_calibration_option == "Path 2":
                    print("Selection option Path 2")
                    # Placeholder for what has to be done
                elif selected_calibration_option == "Path 3":
                    print("Selection option Path 3")
                    # Placeholder for what has to be done
                elif selected_calibration_option == "Path 4":
                    print("Selection option Path 4")
                    # Placeholder for what has to be done
                else:
                    print("Seelct a option from drop down and then calibrate")


                self.window["-OUTPUT-"].update(calibration_result)

            if event == 'Check_equip_status':
                sg_status = 0 # Change this to actual output when connection with sig-gen is successfully established
                sa_status = 0 # Change this to actual output when connection with spec-an is successfully established
                self.SetLED(self.window, 'SG_status', 'green' if sg_status else 'red', 'Available' if sg_status else 'Not Available')
                self.SetLED(self.window, 'SA_status', 'green' if sa_status else 'red', 'Available' if sa_status else 'Not Available')

            if event == "Add to Test List":
                self.add_to_test_list(value)  # Collect the selected parameters
                self.clear_selections()

            if event == "Delete All Test Cases":
                self.clear_all_combinatons()
                self.selected_combinations = []
                self.test_list.update(values=[])

            if event == "Clear Selections":
                self.clear_selections()

            if event == "View":
                selected_combination_index_str = value.get("-TESTLIST-", [])

                if selected_combination_index_str:
                    selected_combination_index = int(selected_combination_index_str.split(" ")[-1]) - 1

                    self.view_selected_combination(selected_combination_index, self.selected_combinations)
                else:
                    sg.popup("Please select a Test Case to view.", icon=r'logo1_7DK_2.ico')

            if event == "-VIEW_ALL_TEST_CASES-":
                if self.selected_combinations:
                    # Create a list of strings representing each combination for display
                    combination_strings = []
                    for i, combination in enumerate(self.selected_combinations):
                        combination_string = f"Test Case {i + 1}:"
                        for selected_parameters_list in combination:
                            for test_name, params in selected_parameters_list.items():
                                # Get the section name based on the test name
                                section_name = None
                                for section, tests in self.test_data['Tests'].items():
                                    if test_name in [test['name'] for test in tests]:
                                        section_name = section
                                        break

                                if section_name is not None:
                                    combination_string += f"\nTest Name: {section_name} ({test_name})"
                                    for param, value in params.items():
                                        if isinstance(value, list):
                                            combination_string += f"\n{param}: {', '.join(map(str, value))}"
                                        else:
                                            combination_string += f"\n{param}: {value}"
                        combination_string += "\n\n==================================================================" \
                                              "======"
                        combination_strings.append(combination_string)

                    # Update the output text box with all combinations
                    self.window["-OUTPUT-"].update("\n\n".join(combination_strings))
                else:
                    sg.popup("No combinations to display.", title="View All Combinations")

            if event == "Run Selected Test In Testlist":
                selected_combination_index_str = None  # Initialize with a default value
                try:
                    selected_combination_index_str = value["-TESTLIST-"]  # Get the selected combination label from the listbox
                except:
                    self.report_str += "Nothing to run\n"
                    self.window["-OUTPUT-"].update(self.report_str)

                if selected_combination_index_str is not None:
                    # Extract the combination number from the label
                    comment_layout = [
                        [sg.Text("Enter your comment:")],
                        [sg.InputText(key='-COMMENT-')],
                        [sg.Button('OK'), sg.Button('Cancel')]
                    ]

                    window = sg.Window('Comment Input', comment_layout)

                    while True:
                        event, values = window.read()

                        if event == sg.WINDOW_CLOSED or event == 'Cancel':
                            self.user_comment = "NA"
                            break
                        elif event == 'OK':
                            self.user_comment = values['-COMMENT-']
                            print(f"User's Comment: {self.user_comment}")
                            window.close()
                            break

                    sg.popup_auto_close('Test is about to start, wait until it finishes', auto_close_duration=2,
                                        font=('Helvetica', 12, 'bold'))
                    try:
                        selected_combination_index = int(selected_combination_index_str.split(" ")[-1]) - 1
                    except ValueError:
                        selected_combination_index = -1

                    self.window["-OUTPUT-"].update(self.report_str)
                    self.run_selected_combination(selected_combination_index, self.selected_combinations)

            if event == "Delete Test Case":
                selected_combination_index_str = value.get("-TESTLIST-", [])

                # Check if any item is selected in the listbox
                if selected_combination_index_str:
                    try:
                        selected_combination_index = int(selected_combination_index_str.split(" ")[-1]) - 1
                    except ValueError:
                        selected_combination_index = -1

                    if 0 <= selected_combination_index < len(self.selected_combinations):
                        del self.selected_combinations[selected_combination_index]  # Delete the selected combination
                        self.update_test_list()  # Update the listbox
                        self.view_selected_combination(-1, self.selected_combinations)  # Clear the output
                        self.report_str += f"Deleted the test case {selected_combination_index+1}\n"
                        self.window["-OUTPUT-"].update(self.report_str)

                else:
                    sg.popup("Please select a Test case to delete.")

            if event == "Run All Tests":
                self.run_all_combinations()

            if event == 'Open Loss Calibration File':
                layout = [
                    [sg.Text('Select a file')],
                    [sg.Input(), sg.FileBrowse(key="-FILE-")],
                    [sg.Button('Open'), sg.Button('Cancel')]
                ]

                window = sg.Window('Get Cal file', layout)

                while True:
                    event, values = window.read()

                    if event == sg.WINDOW_CLOSED or event == 'Cancel':
                        break
                    elif event == 'Open':
                        file_name = values["-FILE-"]
                        print(f"Open {file_name}")

                        try:
                            # Open the selected file using the default associated application
                            os.startfile(file_name)
                        except Exception as e:
                            print(f"Error opening file: {e}")

                window.close()

            if event == "Pause":
                if self.running_combinations:
                    self.pause_combinations = True
                    self.report_str += "Paused the test\n"
                    self.window["-OUTPUT-"].update(self.report_str)
                else:
                    self.report_str += "Couldn't Pause the test\n"
                    self.window["-OUTPUT-"].update(self.report_str)

            if event == "Resume":
                if self.running_combinations and self.pause_combinations:
                    self.pause_combinations = False
                    self.report_str += "Resumed the test\n"
                    self.window["-OUTPUT-"].update(self.report_str )
                else:
                    self.report_str += "Couldn't Resume the test\n"
                    self.window["-OUTPUT-"].update(self.report_str)

            if event == "Stop":
                if self.running_combinations:
                    self.running_combinations = False
                    combinations_thread.join()
                    sg.popup_auto_close("Background Task Stopped", auto_close_duration=5 )
                    self.report_str +="Stopped the test\n"
                    self.window["-OUTPUT-"].update(self.report_str)
                else:
                    self.report_str +="Couldn't stop the test\n"
                    self.window["-OUTPUT-"].update(self.report_str)

            if event == 'Clear Output' or event == "-CLEAR-":
                self.window['-OUTPUT-'].update('')

            if event == 'Save':
                output_data = self.window['-OUTPUT-'].get()
                output_file, screenshot_file = self.save_data_and_screenshot(output_data, window_title="Automated Test GUI")

            if event == 'DUT Configure':
                self.dut_configure()

            if event == 'About':
                sg.Popup("About",title='About this Program', icon=r'logo1_7DK_2.ico')


    def dut_configure(self):
        layout = [
            [sg.Text('DUT Name', size=(20, 1)), sg.Combo(['Random 1', 'Random 2', 'Random 3'], key='dut_name')],
            [sg.Text('DUT Vendor', size=(20, 1)), sg.Combo(['Random 1', 'Random 2', 'Random 3'], key='dut_vendor')],
            [sg.Text('DUT Rev', size=(20, 1)), sg.InputText(key='dut_rev')],
            [sg.Text('DUT SN', size=(20, 1)), sg.InputText(key='dut_sn')],
            [sg.Text('Operator', size=(20, 1)), sg.Combo(['Operator 1', 'Operator 2'], key='operator')],
            [sg.Button('Submit'), sg.Button('Quit')],
            [sg.Text('', size=(30, 2), key='output')]  # Add a Text element to display the output
        ]

        # Initialize input fields with the stored values using StringVar
        dut_window = sg.Window('DUT Configure', layout, finalize=True)
        dut_window['dut_name'].update(self.dut_name)
        dut_window['dut_vendor'].update(self.dut_vendor)
        dut_window['dut_rev'].update(self.dut_rev)
        dut_window['dut_sn'].update(self.dut_sn)

        while True:
            event, values = dut_window.read()
            if event in (sg.WIN_CLOSED, 'Quit'):
                dut_window.close()  # Close the DUT configuration window
                break
            elif event == "Submit":
                self.dut_name = values['dut_name']
                self.dut_vendor = values['dut_vendor']
                self.dut_rev = values['dut_rev']
                self.dut_sn = values['dut_sn']
                self.operator = values['operator']

                dut_window['output'].update(
                    f'DUT Name: {self.dut_name}, DUT Vendor: {self.dut_vendor}, DUT Rev: {self.dut_rev}, DUT SN: {self.dut_sn}')

                try:
                    with open("dut_info.yaml", "r") as file:
                        data = yaml.safe_load(file)
                except FileNotFoundError:
                    self.report_str += "No dut_info.yaml file found"
                    self.window["-OUTPUT-"].update(self.report_str)
                    data = {}

                # Update the data with the current values
                data["dut_name"] = self.dut_name
                data["dut_vendor"] = self.dut_vendor
                data["dut_rev"] = self.dut_rev
                data["dut_sn"] = self.dut_sn
                data["operator"] = self.operator
                # Write the updated data back to the YAML file
                with open("dut_info.yaml", "w") as file:
                    yaml.dump(data, file)
                self.report_str += "Successfully saved the DUT Information\n"
                self.window["-OUTPUT-"].update(self.report_str)
                break
        dut_window.close()


    def get_window_size(self):
        root = mytk.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        root.destroy()
        return screen_width, screen_height

######################################################################################################################

    def test_1(self, temperatures=None, LOFreq=None,  Carrier = None,  modulations=None): # 'Test 1 '
        return f"Test Test 1\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n Modulations={modulations}\n"

######################################################################################################################

    def test_2(self, temperatures=None, LOFreq=None, power_levels=None, Carrier=None):  # 'Test 2'
        return f"Test Test 2 \n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n Power Level={power_levels}\n{self.report_str}\n"

######################################################################################################################
    def test_3(self, temperatures=None, LOFreq=None, Carrier = None, power_levels=None, modulations=None, CA=None):  # 'Test 3'
        return f"Test Test 3\nTest Parameters:\nTemperature={temperatures}\nLOFreq={LOFreq}\nCarrierFreq={Carrier}\nPower Level={power_levels}\nModulations={modulations}\nCA={CA}\n"

######################################################################################################################
    def test_4(self, temperatures=None, LOFreq=None, Carrier = None,  power_levels=None, modulations=None, CA=None):  # 'Test 4'
        return f"Test Test 4\nTest Parameters::\nTemperature={temperatures}\nLOFreq={LOFreq}\nCarrierFreq={Carrier}\nPower Level={power_levels}\nModulations={modulations}\nCA={CA}\n"

######################################################################################################################
    def test_5(self, temperatures=None, LOFreq=None, power_levels=None, Carrier = None, Option = None):  # 'Test 5'
         return f"Test Test 5\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n"

######################################################################################################################
    def test_6(self, temperatures=None, LOFreq=None, power_levels=None, modulations=None, Carrier = None, Option = None):  #'Test 6'
        return f"Test Test 6 \n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n Modulations={modulations}\n"
######################################################################################################################
    def test_7(self, temperatures=None, LOFreq=None, power_levels = None, modulations=None, Carrier = None):  # 'Test 7'
        return f"Test Test 7 \n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n Modulations={modulations}\n Carrier={Carrier}\n"

######################################################################################################################
    def test_8(self, temperatures=None, LOFreq=None):  # 'Test 8'
        return f"Test Test 8 \n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n"

######################################################################################################################
    def test_9(self, temperatures=None, LOFreq=None):  # 'Test 9'
        return f"Test Test 9\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n"

######################################################################################################################
    def test_10(self, temperatures=None, LOFreq=None):  # 'Test 10'
        return f"Test Test 10\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n"

######################################################################################################################
    def test_11(self, temperatures=None, LOFreq=None, modulations=None ):  # 'Test 11'
        return f"Test Test 11\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n Modulations={modulations}\n"

######################################################################################################################
    def test_12(self, temperatures=None, LOFreq=None):  # 'Test 12'
        return f"Test Test 12\n Result:\n Temperature={temperatures}\n LOFreq={LOFreq}\n"

######################################################################################################################


########################################################### DEFINE EVENT FUNCTIONS ###################################################################
    def save_data_and_screenshot(self, output_data, window_title):
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = os.getcwd()
        with open(file_name, 'w') as file:
            file.write(output_data)
        try:
            window = gw.getWindowsWithTitle(window_title)[0]
        except IndexError:
            return file_name, None

        screenshot = ImageGrab.grab(bbox=(window.left, window.top, window.left + window.width, window.top + window.height))
        screenshot_file = f"GUI_screenshot_{current_time}.png"
        screenshot.save(screenshot_file)

        return file_name, screenshot_file

    def load_combinations(self):
        global saved_combinations

        try:
            with open(os.path.join(combinations_file_path, combinations_file_name), "r") as combinations_file:
                saved_combinations = json.load(combinations_file)

            if saved_combinations:
                self.selected_combinations = saved_combinations
                self.update_test_list()
                self.report_str += "Test Cases loaded successfully!\n"
                self.window["-OUTPUT-"].update(self.report_str)
            else:
                self.report_str += "No Test Cases to load\n"
                self.window["-OUTPUT-"].update(self.report_str)
        except FileNotFoundError:
            self.report_str += "Test Cases storage file not found.\n"
            self.window["-OUTPUT-"].update(self.report_str)
        except json.JSONDecodeError as e:
            self.report_str += f"Error loading Test cases: {e}\n"
            self.window["-OUTPUT-"].update(self.report_str)

    def execute_selected_combination(self, selected_combination):
        test_results = []

        for selected_parameters_list in selected_combination:
            for test_name, params in selected_parameters_list.items():

                # Determine the function name dynamically based on the test name
                function_name = f"test_{test_name.split('_')[1]}"  # Extract the number from the test name
                if hasattr(self, function_name):
                    test_function = getattr(self, function_name)

                    test_result = test_function(**params)  # Call the method as a function
                    test_results.append(test_result)

        return test_results

    test_list_values = []
    test_combinations = []
    test_list = sg.DropDown([], size=(20, 8), key="-TESTLIST-")

    def update_test_list(self):
        test_list_values = [f"Test Case {i + 1}" for i in range(len(self.selected_combinations))]
        self.window["-TESTLIST-"].update(values=test_list_values)

    def clear_all_combinatons(self):
        global selected_combinations
        self.selected_combinations = []  # Clear all selected combinations
        self.update_test_list()  # Update the listbox

        with open(os.path.join(combinations_file_path, combinations_file_name), "w") as combinations_file:
            if not self.selected_combinations:
                # If selected_combinations is empty, write an empty list to the file
                json.dump([], combinations_file)
            else:
                json.dump(self.selected_combinations, combinations_file)
        self.report_str += "Cleared all the existing test cases.\nCreate a new one to proceed.\n"
        self.window["-OUTPUT-"].update(self.report_str)

    def clear_selections(self):
        global selected_parameters_list
        selected_parameters_list = []

        # First unhide the hidden paramaters
        for section_name, tests in self.test_data['Tests'].items():
            for test in tests:
                test_name = test['name']
                available_carriers = test.get('Carrier', [])
                for carrier in [1, 2, 3, 4, 5, 6, 'NA']:
                    if carrier in available_carriers:
                        carrier_key = f"carrier_{test_name}_{carrier}"
                        self.window[carrier_key].update(text=carrier, disabled=False)
        self.window.refresh()

        for section_name, tests in self.test_data['Tests'].items():
            for test in tests:
                test_name = test['name']

                # Clear temperatures checkboxes
                if 'temperatures' in test:
                    for value in test.get('temperatures', []):
                        checkbox_key = f"temp_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)

                # Clear LOFreq checkboxes
                if 'LOFreq' in test:
                    for value in test.get('LOFreq', []):
                        checkbox_key = f"lofreq_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)
                    if 'lf' in test:
                        lofreq_input_key = f"lofreq_input_{test_name}"
                        self.window.find_element(lofreq_input_key).update(value="")

                # Clear Carrier freq checkboxes
                if 'Carrier' in test:
                    for value in test.get('Carrier', []):
                        checkbox_key = f"carrier_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)

                # Clear power_levels checkboxes
                if 'power_levels' in test:
                    for value in test.get('power_levels', []):
                        checkbox_key = f"power_level_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)
                    if 'ui_power' in test:
                        power_input_key = f"power_input_{test_name}"
                        self.window.find_element(power_input_key).update(value="")

                # Clear modulations checkboxes
                if 'modulations' in test:
                    for value in test.get('modulations', []):
                        checkbox_key = f"modulations_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)

                # Clear CA checkboxes
                if 'CA' in test:
                    for value in test.get('CA', []):
                        checkbox_key = f"CA_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)

                # Clear Option checkboxes
                if 'Option' in test:
                    for value in test.get('Option', []):
                        checkbox_key = f"Option_{test_name}_{value}"
                        self.window.find_element(checkbox_key).update(value=False)

        self.report_str+= "Cleared the selected parametrs\n"
        self.window["-OUTPUT-"].update(self.report_str)
        self.window.refresh()

    def view_selected_combination(self, selected_combination_index, selected_combinations):
        if selected_combination_index >= 0 and selected_combination_index < len(self.selected_combinations):
            combination = self.selected_combinations[selected_combination_index]
            output_text = []

            output_text.append(f"Test Case {selected_combination_index + 1}:")

            for selected_parameters_list in combination:
                for test_name, params in selected_parameters_list.items():
                    for section, tests in self.test_data['Tests'].items():
                        if test_name in [test['name'] for test in tests]:
                            section_name = section
                            output_text.append(f"\nTest Name: {section_name} ({test_name})")
                    for param, value in params.items():
                        if isinstance(value, list):
                            output_text.append(f"{param}: {', '.join(map(str, value))}")
                        else:
                            output_text.append(f"{param}: {value}")
                output_text.append("============================================================================")
            self.window["-OUTPUT-"].update("\n".join(output_text))

    selected_combinations = saved_combinations

    def add_to_test_list(self, values):
        # Load existing combinations from a JSON file
        existing_combinations = []
        if os.path.isfile(os.path.join(combinations_file_path, combinations_file_name)):
            with open(os.path.join(combinations_file_path, combinations_file_name), "r") as combinations_file:
                existing_combinations = json.load(combinations_file)

        selected_parameters_list = []  # Create a new list to store selected parameters for each test

        # Iterate through the test sections and collect selected parameters
        for section_name, tests in self.test_data['Tests'].items():
            for test in tests:
                test_name = test['name']
                selected_parameters = {}

                # Iterate through the checkbox types
                for checkbox_type, param_name in self.checkbox_to_param.items():
                    for value in test.get(param_name, []):
                        checkbox_key = f"{checkbox_type}_{test_name}_{value}"
                        if values.get(checkbox_key, False):
                            if param_name == "LOFreq":
                                value *= 1e6
                            selected_parameters[param_name] = selected_parameters.get(param_name, []) + [value]

                power_input_key = f"power_input_{test_name}"
                user_input_value = values.get(power_input_key, "")
                lofreq_input_key = f"lofreq_input_{test_name}"
                user_input_lofreq_value = values.get(lofreq_input_key, "")
                if user_input_value.strip() != "":
                    if "power_levels" in selected_parameters:
                        values_list = [int(val.strip()) for val in user_input_value.split(",")]
                        selected_parameters["power_levels"].extend(values_list)
                    else:
                        values_list = [int(val.strip()) for val in user_input_value.split(",")]
                        selected_parameters["power_levels"] = values_list

                if user_input_lofreq_value.strip() != "":
                    if "LOFreq" in selected_parameters:
                        values_list = [int(val.strip()) for val in user_input_lofreq_value.split(",")]
                        values_list = [x * 1e6 for x in values_list]
                        selected_parameters["LOFreq"].extend(values_list)

                    else:
                        values_list = [int(val.strip()) for val in user_input_lofreq_value.split(",")]
                        values_list = [x * 1e6 for x in values_list]
                        selected_parameters["LOFreq"] = values_list

                if selected_parameters:
                    selected_parameters_list.append(
                        {test_name: selected_parameters})  # Append the selected parameters for this test

        if selected_parameters_list:  # Check if the list is not empty
            # Append new combinations to the existing combinations
            existing_combinations.append(selected_parameters_list)

            self.selected_combinations = existing_combinations  # Update the selected_combinations attribute

            self.update_test_list()  # Update the listbox
            self.report_str += f"Added the test case to test list.\nClick on view to see the added test case.\n\n"
            self.window["-OUTPUT-"].update(self.report_str)

            with open(os.path.join(combinations_file_path, combinations_file_name), "w") as combinations_file:
                json.dump(self.selected_combinations, combinations_file)
        else:
            self.report_str += "No parameters selected for this Test case.\n"
            self.window["-OUTPUT-"].update(self.report_str )

    def run_combinations_thread(self):
        global running_combinations, pause_combinations
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Get the current time
        for index, combination in enumerate(self.selected_combinations):
            if not running_combinations:
                return  # Stop if the flag is set to False
            if pause_combinations:
                while pause_combinations:
                    time.sleep(1)  # Sleep while paused

            test_results = self.execute_selected_combination(combination)
            file_name = f"combination_{index + 1}_results_{current_time}.txt"
            with open(file_name, "w") as result_file:
                result_file.write("\n".join(test_results))

        running_combinations = False  # Mark combinations as finished

    def run_all_combinations_thread(self):
        global running_combinations, pause_combinations
        if running_combinations:
            sg.popup_error("Combinations are already running.")
        else:
            running_combinations = True
            pause_combinations = False
            combinations_thread = threading.Thread(target=self.run_combinations_thread)
            combinations_thread.start()

    def run_selected_combination(self, selected_combination_index, selected_combinations):
        i=0
        self.window['progress'].UpdateBar(0)
        if selected_combination_index >= 0 and selected_combination_index < len(self.selected_combinations):
            combination = self.selected_combinations[selected_combination_index]
            test_results = self.execute_selected_combination(combination)
            # Update the output text box with the results
            self.window["-OUTPUT-"].update("\n".join(test_results))
            self.window['progress'].UpdateBar((i + 1) * 1000 // 1)
            # Save the results to a file
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            section_name = None
            for selected_parameters_list in combination:
                for test_name, params in selected_parameters_list.items():
                    for section, tests in self.test_data['Tests'].items():
                        if test_name in [test['name'] for test in tests]:
                            section_name = section
                            break

            combination_name = f"{section_name}_test case_{selected_combination_index + 1}"
            file_name = f"{self.Data_log_path}/{combination_name}_results_{current_time}.txt"
            self.report_str += f"Test is run and results are saved in {file_name}\n"
            self.window["-OUTPUT-"].update(self.report_str)
            with open(file_name, "w") as result_file:
                result_file.write("\n".join(test_results))

        else:
            # If no combination is selected, show a message
            self.report_str += "No combination selected. Please add a combination.\n"
            self.window["-OUTPUT-"].update(self.report_str)

    def run_all_combinations(self):
        comment_layout = [
            [sg.Text("Enter your comment:")],
            [sg.InputText(key='-COMMENT-')],
            [sg.Button('OK'), sg.Button('Cancel')]
        ]
        window = sg.Window('Comment Input', comment_layout)
        while True:
            event, values = window.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancel':
                self.user_comment = "NA"
                break
            elif event == 'OK':
                self.user_comment = values['-COMMENT-']
                print(f"User's Comment: {self.user_comment}")
                window.close()
                break

        sg.popup_auto_close('ALL Tests is about to start, wait until it finishes', auto_close_duration=2,
                            font=('Helvetica', 12, 'bold'))
        # self.user_comment = ''
        global selected_combinations
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Get the current time

        if not self.selected_combinations:
            self.report_str += "Nothing to run\n"
            self.window["-OUTPUT-"].update(self.report_str)
            return
        #print(self.selected_combinations)
        i = 0
        self.window['progress'].UpdateBar(0)
        for index, combination in enumerate(self.selected_combinations):
            test_results = self.execute_selected_combination(combination)
            self.window['progress'].UpdateBar((i + 1) * 1000 // len(self.selected_combinations))
            i = i+1
            # Get the section name based on the test name in the first combination
            section_name = None
            for selected_parameters_list in combination:
                for test_name, params in selected_parameters_list.items():
                    for section, tests in self.test_data['Tests'].items():
                        if test_name in [test['name'] for test in tests]:
                            section_name = section
                            break

            # Create a unique combination name based on section name, index, and current time
            combination_name = f"test case_{index + 1}_{section_name}"
            file_name = f"{self.Data_log_path}/{combination_name}_results_{current_time}.txt"
            with open(file_name, "w") as result_file:
                result_file.write("\n".join(test_results))
        sg.popup_auto_close("\t\t\tDone Testing!! \nAll test cases have been run and results are saved to separate files.", title='Feedback', text_color="white", font=('Helvetica', 12, 'bold'), auto_close_duration=3)


################################################# UNIT TEST ##############################################
if __name__ == '__main__':

    ag = AutomatedGui()
    ag.mainloop()
