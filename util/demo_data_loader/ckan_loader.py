import json
import logging
import os
import ckanapi
import time


CONFIG_FILENAME = os.getenv('CONFIG_FILENAME', 'config.json')
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)

with open(CONFIG_PATH, 'r') as config_file:
    CONFIG = json.loads(config_file.read())['config']

DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG['data_path'])
USERS_FILE = os.path.join(DATA_PATH, CONFIG['users_file'])
ORGANIZATIONS_FILE = os.path.join(DATA_PATH, CONFIG['organizations_file'])
GROUPS_FILE = os.path.join(DATA_PATH, CONFIG['groups_file'])
DATASETS_FILE = os.path.join(DATA_PATH, CONFIG['datasets_file'])
RESOURCE_FOLDER = os.path.join(DATA_PATH, CONFIG['resource_folder'])

log = logging.getLogger(__name__)
log.setLevel(logging.INFO)


def keep_trying(f):
    """
    Retry an api call if it fails the first time.
    CKAN often throws internal errors when bombarded with too many requests.
    """
    def wrapper_function(*args):
        counter = 0
        while True:
            try:
                result = f(*args)
            except ckanapi.errors.ValidationError as e:
                raise e  # Catch all CKANAPIErrors that are not ValidationErrors
            except ckanapi.errors.CKANAPIError as e:
                if counter > 4:
                    raise e  # Raise error after 5 failed attempts
                log.error(f"CKAN API Error: {args[1]['name']}: {e}")
                log.error("Giving CKAN 5s to fix itself before trying again")
                time.sleep(5)
                counter = counter + 1
            else:
                break
        return result
    return wrapper_function


@keep_trying
def create_user(ckan, user):
    try:
        user = ckan.action.user_create(**user)
        log.info(f"Created user {user['name']}")
    except ckanapi.errors.ValidationError:
        log.warning(f"User {user['name']} might exist. Will try to update.")
        user = update_user(ckan, user)
    return user


@keep_trying
def update_user(ckan, user):
    user_id = ckan.action.user_show(id=user['name'])['id']
    user = ckan.action.user_update(id=user_id, **user)
    log.info(f"Updated user {user['name']}")
    return user


@keep_trying
def create_organization(ckan, organization):
    try:
        organization = ckan.action.organization_create(**organization)
        log.info(f"Created organization {organization['name']}")
    except ckanapi.errors.ValidationError:
        log.warning(f"Organization {organization['name']} might exist. Will try to update.")
        organization = update_organization(ckan, organization)
    return organization


@keep_trying
def update_organization(ckan, organization):
    org_id = ckan.action.organization_show(id=organization['name'])['id']
    organization = ckan.action.organization_update(id=org_id, **organization)
    log.info(f"Updated organization {organization['name']}")
    return organization


@keep_trying
def create_dataset(ckan, dataset):
    try:
        dataset = ckan.action.package_create(**dataset)
        log.info(f"Created dataset {dataset['name']}")
    except ckanapi.errors.ValidationError:
        log.warning(f"Dataset {dataset['name']} might exist. Will try to update.")
        dataset = update_dataset(ckan, dataset)
    return dataset


@keep_trying
def update_dataset(ckan, dataset):
    dataset_id = ckan.action.package_show(id=dataset['name'])['id']
    dataset = ckan.action.package_update(id=dataset_id, **dataset)
    log.info(f"Updated dataset {dataset['name']}")
    return dataset


@keep_trying
def create_resource(ckan, resource):
    file_path = os.path.join(RESOURCE_FOLDER, resource['filename'])
    with open(file_path, 'rb') as res_file:
        resource = ckan.call_action(
            'resource_create',
            resource,
            files={'upload': res_file}
        )
    log.info(f"Created resource {resource['name']}")
    return resource


@keep_trying
def create_group(ckan, group):
    try:
        group = ckan.action.group_create(**group)
        log.info(f"Created group {group['name']}")
    except ckanapi.errors.ValidationError:
        log.warning(f"Group {group['name']} might exist. Will try to update.")
        group = update_group(ckan, group)
    return group


@keep_trying
def update_group(ckan, group):
    group_id = ckan.action.group_show(id=group['name'])['id']
    group = ckan.action.group_update(id=group_id, **group)
    log.info(f"Updated group {group['name']}")
    return group


def load_users(ckan):
    """
    Helper method to load users from USERS_FILE config json file
    :param ckan: ckanapi instance
    :return: None
    """
    with open(USERS_FILE, 'r') as users_file:
        users = json.load(users_file)['users']
        for user in users:
            try:
                create_user(ckan, user)
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create user {user['name']}: {e.error_dict}")


def load_organizations(ckan):
    """
    Helper method to load organizations from the ORGANIZATIONS_FILE config file
    :param ckan: ckanapi instance
    :return: a dictionary map of created organization names to their ids
    """
    with open(ORGANIZATIONS_FILE, 'r') as organizations_file:
        organizations = json.load(organizations_file)['organizations']
        for organization in organizations:
            try:
                organization = create_organization(ckan, organization)
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create organization {organization['name']}: {e.error_dict}")


def load_datasets(ckan):
    """
    Helper method to load datasets from the DATASETS_FILE config file
    :param ckan: ckanapi instance
    :return: None
    """
    with open(DATASETS_FILE, 'r') as datasets_file:
        datasets = json.load(datasets_file)['datasets']
        for dataset in datasets:
            resources = dataset.pop('resources', [])
            dataset['resources'] = []
            try:
                dataset = create_dataset(ckan, dataset)
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create dataset {dataset['name']}: {e.error_dict}")
                continue
            for resource in resources:
                resource['package_id'] = dataset['id']
                try:
                    create_resource(ckan, resource)
                except ckanapi.errors.ValidationError as e:
                    log.error(f"Can't create resource {resource['name']}: {e.error_dict}")
                    break


def load_groups(ckan):
    """
    Helper method to load groups from the GROUPS_FILE config file
    :param ckan: ckanapi instance
    :return: None
    """
    with open(GROUPS_FILE, 'r') as groups_file:
        groups = json.load(groups_file)['groups']
        for group in groups:
            try:
                group = create_group(ckan, group)
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create group {group['name']}: {e.error_dict}")


def load_data(ckan_url, ckan_api_key):
    ckan = ckanapi.RemoteCKAN(ckan_url, apikey=ckan_api_key)
    load_users(ckan)
    load_organizations(ckan)
    load_groups(ckan)
    load_datasets(ckan)


if __name__ == '__main__':
    try:
        assert CONFIG['ckan_api_key'] != ''
        load_data(ckan_url=CONFIG['ckan_url'], ckan_api_key=CONFIG['ckan_api_key'])
    except AssertionError:
        log.error('CKAN api key missing from config.json')
