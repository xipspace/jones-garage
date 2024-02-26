import os
import json
import re

selected_file = None

def list_json_files():
    json_files = [file for file in os.listdir() if file.endswith('.json')]
    if not json_files:
        print("No JSON files found in the current directory.")
    else:
        print("JSON Files:")
        for i, file in enumerate(json_files, 1):
            print(f"{i}. {file}")
    return json_files

def select_json_file(json_files):
    global selected_file
    if not json_files:
        print("No JSON files found.")
        return None
    while True:
        try:
            choice = int(input("Enter the number of the JSON file you want to select (0 to go back to the menu): "))
            if choice == 0:
                selected_file = None
                return None
            elif 1 <= choice <= len(json_files):
                selected_file = json_files[choice - 1]
                with open(selected_file, 'r') as f:
                    data = json.load(f)
                num_objects = len(data)
                return selected_file, data, num_objects
            else:
                print("Invalid choice. Please enter a valid number.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def merge_operations(selected_file, data):
    updated_count = 0

    for obj_key, obj_value in data.items():
        if 'properties' in obj_value:
            properties = obj_value['properties']
            ongoing = properties.get('ongoing', False)
            completed = properties.get('completed', False)

            operations = 'none'
            if ongoing and completed:
                operations = 'both'
            elif ongoing:
                operations = 'ongoing'
            elif completed:
                operations = 'completed'

            # Insert 'operations' before 'ongoing' and 'completed' if they exist
            new_properties = {}
            for key, value in properties.items():
                if key == 'ongoing':
                    new_properties['operations'] = operations
                new_properties[key] = value

            obj_value['properties'] = new_properties

            updated_count += 1

    remove_ongoing_completed_keys(data)

    with open(selected_file, 'w') as f:
        json.dump(data, f, indent=4)

    return updated_count

def remove_ongoing_completed_keys(data):
    for obj_key, obj_value in data.items():
        if 'properties' in obj_value:
            properties = obj_value['properties']
            if 'ongoing' in properties:
                del properties['ongoing']
            if 'completed' in properties:
                del properties['completed']

def update_noc_values(selected_file, data):
    updated_count = 0

    for obj_key, obj_value in data.items():
        if 'label' in obj_value:
            label = obj_value['label'].lower()
            match = re.search(r'(\d+)\s*days', label)
            if match:
                days = match.group(1)
                properties = obj_value.get('properties', {})
                properties['noc'] = days
                updated_count += 1
            elif 'days' in label:
                properties = obj_value.get('properties', {})
                properties['noc'] = 'invalid'
                updated_count += 1
            else:
                properties = obj_value.get('properties', {})
                properties['noc'] = 'none'
                updated_count += 1

    with open(selected_file, 'w') as f:
        json.dump(data, f, indent=4)

    return updated_count

def fix_blanket_value(selected_file, data):
    updated_count = 0

    for obj_key, obj_value in data.items():
        label = obj_value.get('label', '').lower()
        if ('blanket additional insured' in label or 'blanket ai' in label) and obj_value.get('properties'):
            properties = obj_value['properties']
            if 'blanket' in properties and not properties['blanket']:
                properties['blanket'] = True
                updated_count += 1

    with open(selected_file, 'w') as f:
        json.dump(data, f, indent=4)

    return updated_count

def update_aggregates(selected_file, data):
    count_both = 0
    count_project = 0
    count_location = 0

    for obj_key, obj_value in data.items():
        if 'label' in obj_value:
            label = obj_value['label'].lower()
            match_project = 'project' in label
            match_location = 'location' in label

            # Update 'agg' property based on keyword matches
            properties = obj_value.get('properties', {})
            if match_project and match_location:
                properties['agg'] = 'both'
                count_both += 1
            elif match_project:
                properties['agg'] = 'project'
                count_project += 1
            elif match_location:
                properties['agg'] = 'location'
                count_location += 1
            else:
                properties['agg'] = 'none'

            obj_value['properties'] = properties

    # Print statistics
    print("\nStatistics:")
    print(f"Both: {count_both}")
    print(f"Project: {count_project}")
    print(f"Location: {count_location}")

    # Save the JSON file
    with open(selected_file, 'w') as f:
        json.dump(data, f, indent=4)

def clear_none(selected_file, data):
    updated_count = 0

    for obj_key, obj_value in data.items():
        if 'properties' in obj_value:
            properties = obj_value['properties']
            keys_to_remove = [key for key, value in properties.items() if value == 'none']
            for key in keys_to_remove:
                del properties[key]
                updated_count += 1

    # Save the JSON file
    with open(selected_file, 'w') as f:
        json.dump(data, f, indent=4)

    print(f"{updated_count} objects were updated.")

def main():
    while True:
        print("\nMenu:")
        print("1. List JSON files")
        print("2. Select JSON file")
        print("0. Quit")
        option = input("Enter your choice: ")

        if option == '1':
            list_json_files()
        elif option == '2':
            json_files = list_json_files()
            if json_files:
                selected_file_info = select_json_file(json_files)
                if selected_file_info:
                    selected_file, data, num_objects = selected_file_info
                    print(f"You selected: {selected_file}")
                    print(f"Number of objects in the file: {num_objects}")
        elif option == '0':
            print("Exiting the script.")
            break
        else:
            print("Invalid option. Please try again.")

        if option == '2':
            while True:
                print("\nJSON Operations Menu:")
                print("1. Merge Operations")
                print("2. Update cancellation days")
                print("3. Fix Blanket Value")
                print("4. Update Aggregates")
                print("9. Clear None")
                print("0. Back to Main Menu")
                sub_option = input("Enter your choice: ")

                if sub_option == '1':
                    if selected_file_info:
                        updated_count = merge_operations(selected_file, data)
                        print(f"{updated_count} objects were updated.")
                elif sub_option == '2':
                    if selected_file_info:
                        updated_count = update_noc_values(selected_file, data)
                        print(f"{updated_count} objects were updated.")
                elif sub_option == '3':
                    if selected_file_info:
                        updated_count = fix_blanket_value(selected_file, data)
                        print(f"{updated_count} objects were updated.")
                elif sub_option == '4':
                    if selected_file_info:
                        update_aggregates(selected_file, data)
                        print("Aggregates updated.")
                elif sub_option == '9':
                    if selected_file_info:
                        clear_none(selected_file, data)
                elif sub_option == '0':
                    print("Returning to the main menu.")
                    break
                else:
                    print("Invalid option. Please try again.")

if __name__ == "__main__":
    main()