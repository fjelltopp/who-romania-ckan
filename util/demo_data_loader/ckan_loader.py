import json
import logging
import os
import ckanapi


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
                ckan.action.user_create(**user)
                log.info(f"Created user {user['name']}")
                continue
            except ckanapi.errors.ValidationError:
                pass  # fallback to user update
            try:
                log.warning(f"User {user['name']} might exist. Will try to update.")
                id = ckan.action.user_show(id=user['name'])['id']
                ckan.action.user_update(id=id, **user)
                log.info(f"Updated user {user['name']}")
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create user {user['name']}: {e.error_dict}")


def load_organizations(ckan):
    """
    Helper method to load organizations from the ORGANIZATIONS_FILE config file
    :param ckan: ckanapi instance
    :return: a dictionary map of created organization names to their ids
    """
    organization_ids_dict = {}
    with open(ORGANIZATIONS_FILE, 'r') as organizations_file:
        organizations = json.load(organizations_file)['organizations']
        for organization in organizations:
            org_name = organization['name']
            try:
                org = ckan.action.organization_create(**organization)
                log.info(f"Created organization {org_name}")
                organization_ids_dict[org_name] = org["id"]
                continue
            except ckanapi.errors.ValidationError:
                pass  # fallback to organization update
            try:
                log.warning(f"Organization {org_name} might exist. Will try to update.")
                org_id = ckan.action.organization_show(id=org_name)['id']
                ckan.action.organization_update(id=org_id, **organization)
                organization_ids_dict[org_name] = org_id
                log.info(f"Updated organization {org_name}")
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create organization {org_name}: {e.error_dict}")
    return organization_ids_dict


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
                dataset = ckan.action.package_create(**dataset)
                log.info(f"Created dataset {dataset['name']}")
            except ckanapi.errors.ValidationError:
                try:
                    log.warning(f"Dataset {dataset['name']} might exist. Will try to update.")
                    id = ckan.action.package_show(id=dataset['name'])['id']
                    dataset = ckan.action.package_update(id=id, **dataset)
                    log.info(f"Updated dataset {dataset['name']}")
                except ckanapi.errors.ValidationError as e:
                    log.error(f"Can't create dataset {dataset['name']}: {e.error_dict}")
                    continue
            for resource in resources:
                file_path = os.path.join(RESOURCE_FOLDER, resource['filename'])
                resource['package_id'] = dataset['id']
                try:
                    with open(file_path, 'rb') as res_file:
                        resource = ckan.call_action(
                            'resource_create',
                            resource,
                            files={'upload': res_file}
                        )
                    log.info(f"Created resource {resource['name']}")
                except ckanapi.errors.ValidationError as e:
                    log.error(f"Can't create resource {resource['name']}: {e.error_dict}")


def load_groups(ckan):
    """
    Helper method to load groups from the GROUPS_FILE config file
    :param ckan: ckanapi instance
    :return: None
    """
    group_ids_dict = {}

    with open(GROUPS_FILE, 'r') as groups_file:
        groups = json.load(groups_file)['groups']

        for group in groups:
            group_name = group['name']
            try:
                org = ckan.action.group_create(**group)
                log.info(f"Created group {group_name}")
                group_ids_dict[group_name] = org["id"]
                continue
            except ckanapi.errors.ValidationError:
                pass  # fallback to group update
            try:
                log.warning(f"Group {group_name} might exist. Will try to update.")
                group_id = ckan.action.group_show(id=group_name)['id']
                ckan.action.group_update(id=group_id, **group)
                group_ids_dict[group_name] = group_id
                log.info(f"Updated group {group_name}")
            except ckanapi.errors.ValidationError as e:
                log.error(f"Can't create group {group_name}: {e.error_dict}")

    return group_ids_dict


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

