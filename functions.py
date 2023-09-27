from user_settings import *

# function for filter logic
def item_filter(item, user_filter, filter_name, filter_negate):
    # if filter is not set, always return true
    if user_filter is None:
        return True
    # if device_list is a string, turn it into list
    if isinstance(user_filter, str):
        local_filter = [user_filter]
    elif isinstance(user_filter, list):
        local_filter = user_filter
    else:
        local_filter = None
        print(f"Error: '{filter_name}' needs to be either a string or a list of strings.")
        quit(1)
    # filter devices
    if item in local_filter:
        result = True
    else:
        result = False
    # return result
    return result if filter_negate is False else not result


# function for device filter
def device_filter(device):
    return item_filter(device, device_list, 'device_list', negate_devices)

# function for variable filter
def variable_filter(variable):
    return item_filter(variable, variable_list, 'variable_list', negate_variables)

