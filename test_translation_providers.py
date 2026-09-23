import hashlib
import json
import queue
import tempfile
import threading
import time
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

import translation_gui
from translation_gui import (
    API_BATCH_LIMITS,
    API_PROVIDERS,
    TranslationApp,
    TranslationProviderError,
    TranslationResult,
    default_provider_configs,
)


class FakeResponse:
    def __init__(self, payload, status_code=200, text=''):
        self._payload = payload
        self.status_code = status_code
        self.ok = status_code < 400
        self.text = text

    def json(self):
        return self._payload

    def raise_for_status(self):
        if not self.ok:
            raise RuntimeError(f'HTTP {self.status_code}')


class FakeSession:
    def __init__(self, response):
        self.response = response
        self.requests = []
        self.headers = {}
        self.trust_env = True

    def request(self, method, url, **kwargs):
        self.requests.append((method, url, kwargs))
        return self.response

    def post(self, url, **kwargs):
        self.requests.append(('POST', url, kwargs))
        return self.response


class FakeWidget:
    def __init__(self):
        self.options = {}
        self.value = None

    def config(self, **kwargs):
        self.options.update(kwargs)

    def set(self, value):
        self.value = value


class FakeListbox:
    def __init__(self):
        self.delete_calls = []
        self.insert_calls = []

    def delete(self, *args):
        self.delete_calls.append(args)

    def insert(self, *args):
        self.insert_calls.append(args)


class FakeValue:
    def __init__(self, value=''):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeMenu:
    def __init__(self):
        self.popup_calls = []
        self.release_calls = 0

    def tk_popup(self, x_root, y_root):
        self.popup_calls.append((x_root, y_root))

    def grab_release(self):
        self.release_calls += 1


class FakeRightClickEvent:
    def __init__(self, x_root, y_root):
        self.x_root = x_root
        self.y_root = y_root


def build_app(response):
    app = TranslationApp.__new__(TranslationApp)
    app.provider_configs = default_provider_configs()
    app.translation_cache = {}
    app._http_local = threading.local()
    app._http_local.session = FakeSession(response)
    app._log_queue = queue.Queue()
    app._provider_rate_lock = threading.Lock()
    app._provider_next_request_at = {}
    app.log_messages = []
    app.log = app.log_messages.append
    app.api_type = 'google'
    return app


def build_batch_workflow_app(provider, item_count):
    app = build_app(FakeResponse({}))
    app.source_files = ['input.xml']
    app.source_folders = []
    app.headless = True
    app.selected_langs = ['USA']
    app.api_type = provider
    app.translation_table = {
        f'测试文本{index}': {'USA': ''}
        for index in range(item_count)
    }
    app.translation_complete = False
    app.translation_failures = []
    app.start_btn = FakeWidget()
    app.confirm_btn = FakeWidget()
    app.progress_var = FakeWidget()
    app.progress_label = FakeWidget()
    app.extract_chinese_texts = lambda: set(app.translation_table)
    app.save_translation_table = lambda: None
    app.abbreviate_translations = lambda _texts=None: None
    if provider == 'tencent':
        app.provider_configs[provider]['secret_id'] = 'test-id'
        app.provider_configs[provider]['secret_key'] = 'test-secret'
    elif provider in ('deepl_free', 'deepl_pro'):
        app.provider_configs[provider]['auth_key'] = 'test-key'
    elif provider == 'volcengine':
        app.provider_configs[provider]['access_key'] = 'test-id'
        app.provider_configs[provider]['secret_key'] = 'test-secret'
    return app


def wait_for_translation(app):
    deadline = time.monotonic() + 5
    while not app.translation_complete and time.monotonic() < deadline:
        time.sleep(0.01)


def resmap_xml(content):
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        f'<ResMap><Message ID="1" Content="{content}" /></ResMap>'
    )


def build_mb_backup_app():
    app = TranslationApp.__new__(TranslationApp)
    app.file_listbox = FakeListbox()
    app.headless = True
    app.logger = None
    app.log_messages = []
    app.log = app.log_messages.append
    app.source_files = []
    app.source_folder = ''
    app.source_folders = []
    app.output_dir = ''
    app.diskc_mode = False
    app.diskc_root = ''
    app.diskc_res_source = ''
    app.diskc_xml_map = {}
    app.diskc_plugin_source = ''
    app.mb_backup_mode = False
    app.mb_backup_archive_path = ''
    app.mb_backup_input_type = ''
    app.mb_backup_root_path = ''
    app.mb_backup_work_dir = ''
    app.mb_backup_sources = {}
    app.mb_backup_output_base_path = ''
    app.mb_backup_output_path = ''
    app.selected_langs = []
    app.translation_table = {}
    return app


class TranslationProviderTests(unittest.TestCase):
    def test_workflow_menu_posts_and_saves_the_selected_choice(self):
        self.assertEqual(
            translation_gui.WORKFLOW_SELECTION_OPTIONS,
            (
                ('marking_cam', '打标Cam工作流'),
                ('mb_backup', 'MB备份工作流'),
                ('simulator', '模拟器工作流'),
            )
        )

        app = TranslationApp.__new__(TranslationApp)
        app.workflow_menu = FakeMenu()
        app.workflow_choice_var = FakeValue('mb_backup')
        app.selected_workflow = ''
        app.log_messages = []
        app.log = app.log_messages.append
        save_calls = []
        app.save_config = lambda: save_calls.append(True)

        self.assertEqual(
            app._show_workflow_menu(FakeRightClickEvent(120, 240)),
            'break'
        )
        app._save_selected_workflow()

        self.assertEqual(app.workflow_menu.popup_calls, [(120, 240)])
        self.assertEqual(app.workflow_menu.release_calls, 1)
        self.assertEqual(app.selected_workflow, 'mb_backup')
        self.assertEqual(save_calls, [True])
        self.assertIn('MB备份工作流', app.log_messages[0])
        self.assertNotIn('后端尚未实现', app.log_messages[0])

    def test_mb_input_preference_is_loaded_and_saved_locally(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps({'mb_backup_source_type': 'folder'}),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.default_maximized = False
            app.selected_workflow = ''
            app.mb_backup_source_type = 'zip'
            app.pack_res_default = False
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log = lambda _message: None
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
                self.assertEqual(app.mb_backup_source_type, 'folder')
                self.assertFalse(app.pack_res_default)
                app.mb_backup_source_type = 'zip'
                app.pack_res_default = True
                app.save_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertEqual(saved_config['mb_backup_source_type'], 'zip')
            self.assertTrue(saved_config['pack_res_default'])

    def test_mb_backup_picker_uses_saved_input_preference_without_prompt(self):
        app = TranslationApp.__new__(TranslationApp)
        app.mb_backup_source_type = 'folder'
        with patch('translation_gui.filedialog.askdirectory', return_value='C:\\MB') as askdirectory:
            with patch('translation_gui.messagebox.askyesnocancel') as prompt:
                self.assertEqual(app._choose_mb_backup_input(), 'C:\\MB')
        askdirectory.assert_called_once_with(
            title='选择已解压的MB备份文件夹',
            mustexist=True
        )
        prompt.assert_not_called()

        app.mb_backup_source_type = 'zip'
        with patch(
            'translation_gui.filedialog.askopenfilename',
            return_value='C:\\MB.zip'
        ) as askopenfilename:
            with patch('translation_gui.messagebox.askyesnocancel') as prompt:
                self.assertEqual(app._choose_mb_backup_input(), 'C:\\MB.zip')
        askopenfilename.assert_called_once_with(
            title='选择MB备份ZIP文件',
            filetypes=[('MB备份 ZIP文件', '*.zip'), ('ZIP文件', '*.zip')]
        )
        prompt.assert_not_called()

    def test_marking_cam_selection_enables_res_packing_and_dispatches_its_backend(self):
        app = TranslationApp.__new__(TranslationApp)
        app.workflow_choice_var = FakeValue('marking_cam')
        app.selected_workflow = ''
        app.pack_var = FakeValue(False)
        app.diskc_workflow_button = FakeWidget()
        app.log_messages = []
        app.log = app.log_messages.append
        app.save_config = lambda: None

        app._save_selected_workflow()
        backend_calls = []
        app.setup_marking_cam_sources = lambda: backend_calls.append(True)
        app._start_selected_workflow()

        self.assertEqual(app.selected_workflow, 'marking_cam')
        self.assertTrue(app.pack_var.get())
        self.assertIn('适用于打标/玻切DiskC文件夹翻译', app.log_messages[0])
        self.assertEqual(
            app.diskc_workflow_button.options['text'],
            '打标Cam工作流'
        )
        self.assertEqual(backend_calls, [True])

    def test_mb_backup_and_simulator_default_to_no_res_packing(self):
        for workflow_id in ('mb_backup', 'simulator'):
            app = TranslationApp.__new__(TranslationApp)
            app.selected_workflow = workflow_id
            app.pack_res_default = True
            app.pack_var = FakeValue(True)

            app._apply_workflow_defaults()

            self.assertFalse(app.pack_res_default)
            self.assertFalse(app.pack_var.get())

    def test_marking_cam_backend_reuses_diskc_source_setup(self):
        app = TranslationApp.__new__(TranslationApp)
        app.log_messages = []
        app.log = app.log_messages.append
        setup_calls = []
        app.setup_diskc_sources = lambda root=None: setup_calls.append(root)

        app.setup_marking_cam_sources('C:\\MarkingCam')

        self.assertEqual(setup_calls, ['C:\\MarkingCam'])
        self.assertIn('复用 DiskC 工作流后端', app.log_messages[0])

    def test_mb_backup_workflow_stages_selected_resources_and_rebuilds_one_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            archive_path = Path(directory) / 'M3R15152_20260526_MBL_297067DA.zip'
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as archive:
                archive.comment = b'mb-backup-comment'
                archive.writestr(
                    'AlarmMacro/AlarmMacro_CHS.xml',
                    resmap_xml('测试')
                )
                archive.writestr(
                    'OCRes/CHS/String/nested/Main.xml',
                    resmap_xml('测试')
                )
                archive.writestr(
                    'OCRes/CHT/String/nested/Main.xml',
                    resmap_xml('旧翻译')
                )
                archive.writestr('BackupOnly/keep.bin', b'keep-this-data')

            app = build_mb_backup_app()
            app.setup_mb_backup_sources(str(archive_path))

            self.assertTrue(app.mb_backup_mode)
            self.assertEqual(len(app.source_files), 2)
            self.assertEqual(
                {source.category for source in app.mb_backup_sources.values()},
                {'alarm_macro', 'ocres_string'}
            )
            self.assertTrue(
                any('Ladder/AlarmPLC_CHS.xml' in message for message in app.log_messages)
            )
            self.assertTrue(
                any('ParameterExt/ParamExt_CHS.xml' in message for message in app.log_messages)
            )
            self.assertTrue(
                any('ParameterExt/ParamExt_RBit_CHS.xml' in message for message in app.log_messages)
            )

            app.selected_langs = ['ENG', 'CHT']
            app.translation_table = {'测试': {'ENG': 'Test', 'CHT': '測試'}}
            output_path = app.generate_mb_backup_file()

            self.assertTrue(Path(output_path).is_file())
            self.assertRegex(
                Path(output_path).name,
                r'^M3R15152_20260526_MBL_[0-9A-F]{8}\.zip$'
            )
            self.assertNotEqual(Path(output_path), archive_path)
            self.assertEqual(app.mb_backup_work_dir, '')

            with zipfile.ZipFile(output_path) as archive:
                self.assertEqual(archive.comment, b'mb-backup-comment')
                self.assertEqual(archive.read('BackupOnly/keep.bin'), b'keep-this-data')
                self.assertEqual(
                    archive.read('AlarmMacro/AlarmMacro_CHS.xml').decode('utf-8'),
                    resmap_xml('测试')
                )
                self.assertEqual(len(archive.namelist()), len(set(archive.namelist())))

                for language, translation in (('ENG', 'Test'), ('CHT', '測試')):
                    for target_entry in (
                        f'AlarmMacro/AlarmMacro_{language}.xml',
                        f'OCRes/{language}/String/nested/Main.xml',
                    ):
                        root = ET.fromstring(archive.read(target_entry))
                        self.assertEqual(root.find('Message').get('Content'), translation)

            self.assertEqual(
                f'{TranslationApp._file_crc32(output_path):08X}',
                Path(output_path).stem.rsplit('_', 1)[1]
            )

    def test_mb_backup_selection_dispatches_to_its_backend(self):
        app = TranslationApp.__new__(TranslationApp)
        app.selected_workflow = 'mb_backup'
        app.log = lambda _message: None
        backend_calls = []
        app.setup_mb_backup_sources = lambda: backend_calls.append(True)

        app._start_selected_workflow()

        self.assertEqual(backend_calls, [True])

    def test_simulator_selection_dispatches_to_its_backend(self):
        app = TranslationApp.__new__(TranslationApp)
        app.selected_workflow = 'simulator'
        app.log = lambda _message: None
        backend_calls = []
        app.setup_simulator_sources = lambda: backend_calls.append(True)

        app._start_selected_workflow()

        self.assertEqual(backend_calls, [True])

    def test_simulator_workflow_translates_only_chs_string_xml(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'DISKC_V1.1.6'
            string_dir = (
                root / 'DiskC' / 'OpenCnc Shared' / 'OCRes' / 'CHS' / 'String'
            )
            string_dir.mkdir(parents=True)
            source_xml = string_dir / 'nested' / 'Main.xml'
            source_xml.parent.mkdir()
            source_xml.write_text(resmap_xml('测试'), encoding='utf-8')
            unrelated_xml = (
                root / 'DiskC' / 'OpenCnc Shared' / 'OCRes' / 'CHS' / 'Other.xml'
            )
            unrelated_xml.write_text(resmap_xml('不应处理'), encoding='utf-8')

            app = TranslationApp.__new__(TranslationApp)
            app.file_listbox = FakeListbox()
            app.headless = True
            app.logger = None
            app.log_messages = []
            app.log = app.log_messages.append
            app.source_files = []
            app.source_folders = []
            app.source_folder = ''
            app.output_dir = ''
            app.diskc_mode = False
            app.diskc_root = ''
            app.diskc_res_source = ''
            app.diskc_xml_map = {}
            app.diskc_plugin_source = ''
            app.simulator_mode = False
            app.simulator_root = ''
            app.simulator_xml_map = {}
            app.mb_backup_mode = False
            app.mb_backup_archive_path = ''
            app.mb_backup_sources = {}
            app.selected_langs = ['ENG', 'CHT']
            app.translation_table = {'测试': {'ENG': 'Test', 'CHT': '測試'}}
            app.pack_var = FakeValue(False)

            app.setup_simulator_sources(str(root))

            self.assertTrue(app.simulator_mode)
            self.assertEqual(
                app.simulator_root,
                str(root / 'DiskC')
            )
            self.assertEqual(len(app.source_files), 1)
            self.assertEqual(
                app.simulator_xml_map[app.source_files[0]],
                'nested\\Main.xml'
            )

            with patch('translation_gui.messagebox.showinfo'):
                app.generate_xml_files()

            for language, translation in (('ENG', 'Test'), ('CHT', '測試')):
                output = (
                    root / 'DiskC' / 'OpenCnc Shared' / 'OCRes' /
                    language / 'String' / 'nested' / 'Main.xml'
                )
                parsed = ET.fromstring(output.read_text(encoding='utf-8'))
                self.assertEqual(
                    parsed.find('Message').get('Content'),
                    translation
                )
            self.assertFalse(
                (root / 'DiskC' / 'OpenCnc Shared' / 'OCRes' /
                 'ENG' / 'Other.xml').exists()
            )

    def test_mb_backup_workflow_routes_all_five_resource_categories(self):
        with tempfile.TemporaryDirectory() as directory:
            archive_path = Path(directory) / 'MGT0001_20260730_MBL_11A2626D.zip'
            source_entries = (
                'AlarmMacro/AlarmMacro_CHS.xml',
                'Ladder/AlarmPLC_CHS.xml',
                'ParameterExt/ParamExt_CHS.xml',
                'ParameterExt/ParamExt_RBit_CHS.xml',
                'OCRes/CHS/String/nested/Main.xml',
            )
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as archive:
                for source_entry in source_entries:
                    archive.writestr(source_entry, resmap_xml('测试'))

            app = build_mb_backup_app()
            app.setup_mb_backup_sources(str(archive_path))
            app.selected_langs = ['ENG']
            app.translation_table = {'测试': {'ENG': 'Test'}}
            output_path = app.generate_mb_backup_file()

            with zipfile.ZipFile(output_path) as archive:
                self.assertEqual(
                    {
                        'AlarmMacro/AlarmMacro_ENG.xml',
                        'Ladder/AlarmPLC_ENG.xml',
                        'ParameterExt/ParamExt_ENG.xml',
                        'ParameterExt/ParamExt_RBit_ENG.xml',
                        'OCRes/ENG/String/nested/Main.xml',
                    },
                    {
                        name for name in archive.namelist()
                        if name.endswith('_ENG.xml')
                        or name.startswith('OCRes/ENG/String/')
                    }
                )

    def test_mb_backup_workflow_accepts_extracted_backup_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            backup_root = Path(directory) / 'M3R15152_20260526_MBL_297067DA'
            source_path = backup_root / 'AlarmMacro' / 'AlarmMacro_CHS.xml'
            source_path.parent.mkdir(parents=True)
            source_path.write_text(resmap_xml('测试'), encoding='utf-8')
            existing_target = backup_root / 'AlarmMacro' / 'AlarmMacro_CHT.xml'
            existing_target.write_text(resmap_xml('旧翻译'), encoding='utf-8')
            unrelated_path = backup_root / 'BackupOnly' / 'keep.bin'
            unrelated_path.parent.mkdir(parents=True)
            unrelated_path.write_bytes(b'keep-this-data')

            app = build_mb_backup_app()
            app.setup_mb_backup_sources(str(backup_root))

            self.assertTrue(app.mb_backup_mode)
            self.assertEqual(app.mb_backup_input_type, 'folder')
            self.assertEqual(app.mb_backup_root_path, str(backup_root))
            self.assertTrue(Path(app.mb_backup_work_dir).is_dir())
            self.assertEqual(len(app.source_files), 1)
            staged_source = next(iter(app.mb_backup_sources.values()))
            self.assertEqual(
                staged_source.archive_entry,
                'AlarmMacro/AlarmMacro_CHS.xml'
            )
            self.assertTrue(Path(staged_source.staged_path).is_file())

            app.selected_langs = ['ENG', 'CHT']
            app.translation_table = {'测试': {'ENG': 'Test', 'CHT': '測試'}}
            output_path = app.generate_mb_backup_file()

            self.assertTrue(Path(output_path).is_file())
            self.assertEqual(app.mb_backup_work_dir, '')
            self.assertRegex(
                Path(output_path).name,
                r'^M3R15152_20260526_MBL_[0-9A-F]{8}\.zip$'
            )
            self.assertEqual(
                existing_target.read_text(encoding='utf-8'),
                resmap_xml('旧翻译')
            )

            with zipfile.ZipFile(output_path) as archive:
                self.assertEqual(archive.read('BackupOnly/keep.bin'), b'keep-this-data')
                for language, translation in (('ENG', 'Test'), ('CHT', '測試')):
                    root = ET.fromstring(
                        archive.read(f'AlarmMacro/AlarmMacro_{language}.xml')
                    )
                    self.assertEqual(root.find('Message').get('Content'), translation)

    def test_selected_workflow_is_loaded_and_saved(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps({'selected_workflow': 'marking_cam'}),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.default_maximized = False
            app.selected_workflow = ''
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log = lambda _message: None
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
                self.assertEqual(app.selected_workflow, 'marking_cam')
                app.selected_workflow = 'simulator'
                app.save_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertEqual(saved_config['selected_workflow'], 'simulator')

            config_path.write_text(
                json.dumps({'selected_workflow': 'unsupported_workflow'}),
                encoding='utf-8'
            )
            app.selected_workflow = 'marking_cam'
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            self.assertEqual(app.selected_workflow, '')
            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertEqual(saved_config['selected_workflow'], '')

    def test_clear_files_resets_diskc_workflow_state(self):
        app = TranslationApp.__new__(TranslationApp)
        app.file_listbox = FakeListbox()
        app.source_files = ['C:\\DiskC\\prepared.xml']
        app.source_folders = ['C:\\DiskC']
        app.source_folder = 'C:\\DiskC'
        app.diskc_mode = True
        app.diskc_root = 'C:\\DiskC'
        app.diskc_res_source = 'C:\\DiskC\\_res_converted\\CHS'
        app.diskc_xml_map = {
            'C:\\DiskC\\String\\CHS.xml': 'CHS.xml',
        }
        app.diskc_plugin_source = 'C:\\DiskC\\Plugin\\Config\\CHS.xml'

        app.clear_files()

        self.assertEqual(app.file_listbox.delete_calls, [(0, translation_gui.tk.END)])
        self.assertEqual(app.source_files, [])
        self.assertEqual(app.source_folders, [])
        self.assertEqual(app.source_folder, '')
        self.assertFalse(app.diskc_mode)
        self.assertEqual(app.diskc_root, '')
        self.assertEqual(app.diskc_res_source, '')
        self.assertEqual(app.diskc_xml_map, {})
        self.assertEqual(app.diskc_plugin_source, '')

    def test_clear_files_deletes_mb_backup_staging_files(self):
        with tempfile.TemporaryDirectory() as directory:
            staging_dir = Path(directory) / 'mb-staging'
            staging_dir.mkdir()
            (staging_dir / 'source.xml').write_text(
                resmap_xml('测试'),
                encoding='utf-8'
            )

            app = TranslationApp.__new__(TranslationApp)
            app.file_listbox = FakeListbox()
            app.source_files = [str(staging_dir / 'source.xml')]
            app.source_folders = []
            app.source_folder = ''
            app.diskc_mode = False
            app.diskc_root = ''
            app.diskc_res_source = ''
            app.diskc_xml_map = {}
            app.diskc_plugin_source = ''
            app.mb_backup_mode = True
            app.mb_backup_archive_path = 'backup.zip'
            app.mb_backup_work_dir = str(staging_dir)
            app.mb_backup_sources = {'source': object()}
            app.mb_backup_output_path = ''
            app.log = lambda _message: None

            app.clear_files()

            self.assertFalse(staging_dir.exists())
            self.assertFalse(app.mb_backup_mode)
            self.assertEqual(app.mb_backup_archive_path, '')
            self.assertEqual(app.mb_backup_work_dir, '')
            self.assertEqual(app.mb_backup_sources, {})

    def test_free_providers_are_listed_first(self):
        self.assertEqual(
            list(API_PROVIDERS)[:3],
            ['deepl_free', 'baidu', 'youdao']
        )
        self.assertNotIn('libretranslate', API_PROVIDERS)

    def test_removed_libretranslate_configuration_falls_back_to_google(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps({
                    'api_type': 'libretranslate',
                    'provider_configs': {
                        'libretranslate': {
                            'endpoint': 'http://127.0.0.1:5000',
                            'api_key': 'obsolete-key',
                        },
                        'baidu': {
                            'app_id': 'existing-app',
                            'secret_key': 'existing-secret',
                        },
                    },
                }),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log_messages = []
            app.log = app.log_messages.append
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertEqual(app.api_type, 'google')
            self.assertNotIn('libretranslate', app.provider_configs)
            self.assertNotIn('libretranslate', saved_config['provider_configs'])
            self.assertEqual(saved_config['provider_configs']['baidu']['app_id'], 'existing-app')
            self.assertIn('LibreTranslate 接口已移除', app.log_messages[0])

    def test_legacy_volcengine_region_is_removed_from_saved_config(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps({
                    'api_type': 'volcengine',
                    'provider_configs': {
                        'volcengine': {
                            'access_key': 'test-access-key',
                            'secret_key': 'test-secret',
                            'region': 'cn-north-1',
                            'endpoint': 'https://translate.volcengineapi.com',
                        },
                    },
                }),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log = lambda _message: None
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertEqual(app.api_type, 'volcengine')
            self.assertEqual(
                saved_config['provider_configs']['volcengine']['endpoint'],
                'https://translate.volcengineapi.com'
            )
            self.assertNotIn('region', saved_config['provider_configs']['volcengine'])

    def test_default_maximized_setting_is_loaded_and_saved(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps({'default_fullscreen': True}),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.default_maximized = False
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log = lambda _message: None
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
                self.assertTrue(app.default_maximized)
                app.default_maximized = False
                app.save_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            saved_config = json.loads(config_path.read_text(encoding='utf-8'))
            self.assertFalse(saved_config['default_maximized'])
            self.assertNotIn('default_fullscreen', saved_config)

    def test_window_mode_can_switch_between_maximized_and_windowed(self):
        class FakeRoot:
            def __init__(self):
                self.state_value = 'normal'
                self.geometry_value = None

            def state(self, value=None):
                if value is None:
                    return self.state_value
                self.state_value = value

            def geometry(self, value):
                self.geometry_value = value

        app = TranslationApp.__new__(TranslationApp)
        app.root = FakeRoot()
        app.log = lambda _message: None

        self.assertTrue(app._apply_window_mode(True))
        self.assertEqual(app.root.state_value, 'zoomed')

        self.assertTrue(app._apply_window_mode(False))
        self.assertEqual(app.root.state_value, 'normal')
        self.assertEqual(app.root.geometry_value, '1280x850')

    def test_deepl_free_uses_the_free_endpoint_and_auth_header(self):
        app = build_app(FakeResponse({'translations': [{'text': 'Hello'}]}))
        app.provider_configs['deepl_free']['auth_key'] = 'test-key'

        result = app._translate_with_provider('deepl_free', '你好', 'en-US')

        self.assertTrue(result.success)
        method, url, kwargs = app._http_local.session.requests[0]
        self.assertEqual(method, 'POST')
        self.assertEqual(url, 'https://api-free.deepl.com/v2/translate')
        self.assertEqual(kwargs['headers']['Authorization'], 'DeepL-Auth-Key test-key')
        self.assertEqual(kwargs['json']['target_lang'], 'EN-US')

    def test_deepl_batch_uses_a_text_array(self):
        app = build_app(FakeResponse({
            'translations': [{'text': 'Hello'}, {'text': 'World'}]
        }))
        app.provider_configs['deepl_free']['auth_key'] = 'test-key'

        results = app._translate_provider_text_batch(
            'deepl_free',
            ['你好', '世界'],
            'en-US'
        )

        self.assertEqual([result.text for result in results], ['Hello', 'World'])
        _, _, kwargs = app._http_local.session.requests[0]
        self.assertEqual(kwargs['json']['text'], ['你好', '世界'])

    def test_ssl_proxy_failure_retries_with_a_direct_session(self):
        class ProxyErrorSession(FakeSession):
            def request(self, method, url, **kwargs):
                self.requests.append((method, url, kwargs))
                raise translation_gui.requests.exceptions.SSLError('proxy TLS EOF')

        app = build_app(FakeResponse({'ok': True}))
        proxy_session = ProxyErrorSession(FakeResponse({}))
        direct_session = FakeSession(FakeResponse({'ok': True}))
        app._http_local.session = proxy_session

        with patch('translation_gui.requests.Session', return_value=direct_session):
            payload = app._request_json(
                'GET',
                'https://api-free.deepl.com/v2/usage'
            )

        self.assertEqual(payload, {'ok': True})
        self.assertEqual(len(proxy_session.requests), 1)
        self.assertEqual(len(direct_session.requests), 1)
        self.assertFalse(direct_session.trust_env)
        self.assertTrue(
            any('自动改用直连请求' in message for message in app.log_messages)
        )

    def test_baidu_signs_the_same_full_text_that_it_sends(self):
        app = build_app(FakeResponse({'trans_result': [{'dst': 'translated'}]}))
        app.provider_configs['baidu']['app_id'] = 'test-app'
        app.provider_configs['baidu']['secret_key'] = 'test-secret'
        text = '这是超过五十个字符的百度翻译测试文本，用于验证签名和请求正文使用同一完整字符串。'

        result = app._translate_with_provider('baidu', text, 'en-US')

        self.assertTrue(result.success)
        _, _, kwargs = app._http_local.session.requests[0]
        self.assertEqual(kwargs['data']['q'], text)

    def test_youdao_uses_v3_request_fields(self):
        app = build_app(FakeResponse({'errorCode': '0', 'translation': ['Hello']}))
        app.provider_configs['youdao']['app_key'] = 'test-app'
        app.provider_configs['youdao']['app_secret'] = 'test-secret'

        result = app._translate_with_provider('youdao', '你好', 'en-US')

        self.assertTrue(result.success)
        _, url, kwargs = app._http_local.session.requests[0]
        self.assertEqual(url, 'https://openapi.youdao.com/api')
        self.assertEqual(kwargs['data']['signType'], 'v3')
        self.assertEqual(kwargs['data']['from'], 'zh-CHS')
        self.assertEqual(kwargs['data']['to'], 'en')

    def test_tencent_uses_tc3_signature(self):
        app = build_app(FakeResponse({'Response': {'TargetText': 'Hello'}}))
        config = app.provider_configs['tencent']
        config['secret_id'] = 'test-id'
        config['secret_key'] = 'test-secret'

        result = app._translate_with_provider('tencent', '你好', 'en-US')

        self.assertTrue(result.success)
        _, _, kwargs = app._http_local.session.requests[0]
        self.assertTrue(kwargs['headers']['Authorization'].startswith('TC3-HMAC-SHA256 Credential=test-id/'))
        self.assertEqual(kwargs['headers']['X-TC-Action'], 'TextTranslate')

    def test_tencent_request_limit_is_retryable(self):
        app = build_app(FakeResponse({
            'Response': {
                'Error': {
                    'Code': 'RequestLimitExceeded',
                    'Message': 'frequency limit exceeded',
                }
            }
        }))
        config = app.provider_configs['tencent']
        config['secret_id'] = 'test-id'
        config['secret_key'] = 'test-secret'

        with self.assertRaises(TranslationProviderError) as context:
            app._translate_tencent('你好', 'en-US', config)

        self.assertTrue(context.exception.retryable)

    def test_tencent_batch_uses_source_and_target_text_lists(self):
        app = build_app(FakeResponse({
            'Response': {'TargetTextList': ['Hello', 'World']}
        }))
        config = app.provider_configs['tencent']
        config['secret_id'] = 'test-id'
        config['secret_key'] = 'test-secret'

        with patch.object(app, '_wait_for_provider_slot') as wait_for_slot:
            results = app._translate_tencent_text_batch(['你好', '世界'], 'en-US')

        self.assertEqual([result.text for result in results], ['Hello', 'World'])
        self.assertTrue(all(result.success for result in results))
        self.assertEqual(wait_for_slot.call_count, 1)
        _, _, kwargs = app._http_local.session.requests[0]
        self.assertEqual(kwargs['headers']['X-TC-Action'], 'TextTranslateBatch')
        payload = json.loads(kwargs['data'].decode('utf-8'))
        self.assertEqual(payload['SourceTextList'], ['你好', '世界'])
        self.assertEqual(payload['Target'], 'en')

    def test_tencent_batch_partition_respects_text_and_length_limits(self):
        app = build_app(FakeResponse({}))
        texts = ['x'] * 201 + ['y' * 2001]

        batches, oversized = app._partition_tencent_texts(texts)

        self.assertEqual([len(batch) for batch in batches], [200, 1])
        self.assertEqual(oversized, ['y' * 2001])

    def test_tencent_batch_retries_a_rate_limit_error(self):
        app = build_app(FakeResponse({}))
        config = app.provider_configs['tencent']
        config['secret_id'] = 'test-id'
        config['secret_key'] = 'test-secret'
        attempts = []

        def rate_limited_then_success(*_args, **_kwargs):
            attempts.append(True)
            if len(attempts) == 1:
                raise TranslationProviderError('RequestLimitExceeded', retryable=True)
            return {'TargetTextList': ['Hello']}

        app._request_tencent_tmt = rate_limited_then_success
        with patch.object(app, '_wait_for_provider_slot') as wait_for_slot, patch('translation_gui.time.sleep'):
            results = app._translate_tencent_text_batch(['你好'], 'en-US')

        self.assertTrue(results[0].success)
        self.assertEqual(results[0].text, 'Hello')
        self.assertEqual(len(attempts), 2)
        self.assertEqual(wait_for_slot.call_count, 2)

    def test_tencent_workflow_groups_201_items_into_two_batch_requests(self):
        app = build_batch_workflow_app('tencent', 201)
        batch_calls = []

        def translate_batch(provider, texts, target_lang):
            batch_calls.append((provider, list(texts), target_lang))
            return [TranslationResult(True, f'Translated {index}') for index, _ in enumerate(texts)]

        app._translate_provider_text_batch = translate_batch
        app.start_translation()
        wait_for_translation(app)

        self.assertTrue(app.translation_complete)
        self.assertEqual(sorted(len(texts) for _, texts, _ in batch_calls), [1, 200])
        self.assertTrue(all(provider == 'tencent' for provider, _, _ in batch_calls))
        self.assertTrue(all(target == 'en-US' for _, _, target in batch_calls))
        self.assertTrue(
            all(values['USA'].startswith('Translated ') for values in app.translation_table.values())
        )
        self.assertEqual(app.translation_failures, [])

    def test_deepl_workflow_groups_51_items_into_two_batch_requests(self):
        app = build_batch_workflow_app('deepl_free', 51)
        batch_calls = []

        def translate_batch(provider, texts, target_lang):
            batch_calls.append((provider, list(texts), target_lang))
            return [TranslationResult(True, f'Translated {index}') for index, _ in enumerate(texts)]

        app._translate_provider_text_batch = translate_batch
        app.start_translation()
        wait_for_translation(app)

        self.assertTrue(app.translation_complete)
        self.assertEqual(sorted(len(texts) for _, texts, _ in batch_calls), [1, 50])
        self.assertTrue(all(provider == 'deepl_free' for provider, _, _ in batch_calls))
        self.assertTrue(all(target == 'en-US' for _, _, target in batch_calls))
        self.assertEqual(app.translation_failures, [])

    def test_translation_workflow_ignores_stale_translation_memory_entries(self):
        app = build_batch_workflow_app('deepl_free', 2)
        app.extract_chinese_texts = lambda: {'测试文本0'}
        batch_calls = []

        def translate_batch(provider, texts, target_lang):
            batch_calls.append((provider, list(texts), target_lang))
            return [TranslationResult(True, f'Translated {text}') for text in texts]

        app._translate_provider_text_batch = translate_batch
        app.start_translation()
        wait_for_translation(app)

        self.assertTrue(app.translation_complete)
        self.assertEqual(
            batch_calls,
            [('deepl_free', ['测试文本0'], 'en-US')]
        )
        self.assertEqual(
            app.translation_table['测试文本0']['USA'],
            'Translated 测试文本0'
        )
        self.assertEqual(app.translation_table['测试文本1']['USA'], '')
        self.assertEqual(app.translation_failures, [])

    def test_translation_workflow_leaves_stale_memory_entries_untouched(self):
        app = build_batch_workflow_app('deepl_free', 3)
        app.extract_chinese_texts = lambda: {'测试文本0'}
        app.abbreviate_translations = TranslationApp.abbreviate_translations.__get__(
            app,
            TranslationApp
        )
        stale_translation = (
            'Historical translation must remain untouched even though it is '
            'much longer than forty characters.'
        )
        app.translation_table['测试文本2']['USA'] = stale_translation
        batch_calls = []

        def translate_batch(provider, texts, target_lang):
            batch_calls.append((provider, list(texts), target_lang))
            return [TranslationResult(True, f'Translated {text}') for text in texts]

        app._translate_provider_text_batch = translate_batch
        app.start_translation()
        wait_for_translation(app)

        self.assertTrue(app.translation_complete)
        self.assertEqual(
            batch_calls,
            [('deepl_free', ['测试文本0'], 'en-US')]
        )
        self.assertEqual(app.translation_table['测试文本1']['USA'], '')
        self.assertEqual(app.translation_table['测试文本2']['USA'], stale_translation)
        self.assertEqual(app.translation_failures, [])

    def test_tencent_rate_limiter_reserves_a_global_request_slot(self):
        app = build_app(FakeResponse({}))

        with patch('translation_gui.time.sleep') as sleep:
            app._wait_for_provider_slot('tencent')
            app._wait_for_provider_slot('tencent')

        sleep.assert_called_once()
        self.assertGreater(sleep.call_args.args[0], 0.20)

    def test_tencent_rate_limit_error_retries_before_returning_a_result(self):
        app = build_app(FakeResponse({}))
        attempts = []

        def rate_limited_then_success(*_args, **_kwargs):
            attempts.append(True)
            if len(attempts) == 1:
                raise TranslationProviderError('RequestLimitExceeded', retryable=True)
            return 'Hello'

        app._translate_provider_once = rate_limited_then_success
        with patch.object(app, '_wait_for_provider_slot') as wait_for_slot, patch('translation_gui.time.sleep'):
            result = app._translate_with_provider('tencent', '你好', 'en-US')

        self.assertTrue(result.success)
        self.assertEqual(result.text, 'Hello')
        self.assertEqual(len(attempts), 2)
        self.assertEqual(wait_for_slot.call_count, 2)

    def test_volcengine_uses_signature_v4_request(self):
        app = build_app(FakeResponse({'TranslationList': [{'Translation': 'Hello'}]}))
        config = app.provider_configs['volcengine']
        config['access_key'] = 'test-access-key'
        config['secret_key'] = 'test-secret'

        result = app._translate_with_provider('volcengine', '你好', 'en-US')

        self.assertTrue(result.success)
        _, url, kwargs = app._http_local.session.requests[0]
        self.assertIn('Action=TranslateText&Version=2020-06-01', url)
        self.assertTrue(kwargs['headers']['Authorization'].startswith('HMAC-SHA256 Credential=test-access-key/'))

    def test_volcengine_signature_matches_official_signer_vector(self):
        app = build_app(FakeResponse({}))
        config = app.provider_configs['volcengine']
        config['access_key'] = 'test-access-key'
        config['secret_key'] = 'test-secret'
        body = b'{"hello":"world"}'

        url, headers, sent_body = app._build_volcengine_signed_request(
            config,
            body,
            timestamp='20260912T054500Z'
        )

        self.assertEqual(url, 'https://translate.volcengineapi.com/?Action=TranslateText&Version=2020-06-01')
        self.assertEqual(sent_body, body)
        self.assertEqual(headers['X-Content-Sha256'], hashlib.sha256(body).hexdigest())
        self.assertEqual(
            headers['Authorization'],
            'HMAC-SHA256 Credential=test-access-key/20260912/cn-beijing/translate/request, '
            'SignedHeaders=content-type;host;x-content-sha256;x-date, '
            'Signature=1c7784a42d6892e3ede2a86c097df15931c3041d27a6c3905f7bac9e2202643c'
        )

    def test_volcengine_batch_uses_text_list(self):
        app = build_app(FakeResponse({
            'TranslationList': [{'Translation': 'Hello'}, {'Translation': 'World'}]
        }))
        config = app.provider_configs['volcengine']
        config['access_key'] = 'test-access-key'
        config['secret_key'] = 'test-secret'

        results = app._translate_provider_text_batch(
            'volcengine',
            ['你好', '世界'],
            'en-US'
        )

        self.assertEqual([result.text for result in results], ['Hello', 'World'])
        _, _, kwargs = app._http_local.session.requests[0]
        payload = json.loads(kwargs['data'].decode('utf-8'))
        self.assertEqual(payload['TextList'], ['你好', '世界'])

    def test_provider_batch_partitions_respect_documented_limits(self):
        app = build_app(FakeResponse({}))

        deepl_batches, deepl_rejected = app._partition_provider_texts(
            'deepl_free',
            ['x'] * (API_BATCH_LIMITS['deepl_free']['max_texts'] + 1)
        )
        volcengine_batches, volcengine_rejected = app._partition_provider_texts(
            'volcengine',
            ['x'] * (API_BATCH_LIMITS['volcengine']['max_texts'] + 1)
        )
        aliyun_batches, aliyun_rejected = app._partition_provider_texts(
            'aliyun',
            ['x'] * (API_BATCH_LIMITS['aliyun']['max_texts'] + 1)
        )
        _, oversized_volcengine = app._partition_provider_texts(
            'volcengine',
            ['x' * 5001]
        )
        _, oversized_aliyun = app._partition_provider_texts(
            'aliyun',
            ['x' * 1001]
        )
        deepl_escaped_batches, deepl_escaped_rejected = app._partition_provider_texts(
            'deepl_free',
            ['你' * 30000]
        )
        volcengine_character_batches, _ = app._partition_provider_texts(
            'volcengine',
            ['x' * 2500, 'x' * 2501]
        )
        aliyun_character_batches, _ = app._partition_provider_texts(
            'aliyun',
            ['x' * 1000] * 9
        )

        self.assertEqual([len(batch) for batch in deepl_batches], [50, 1])
        self.assertEqual(deepl_rejected, [])
        self.assertEqual([len(batch) for batch in volcengine_batches], [16, 1])
        self.assertEqual(volcengine_rejected, [])
        self.assertEqual(len(oversized_volcengine), 1)
        self.assertEqual([len(batch) for batch in aliyun_batches], [50, 1])
        self.assertEqual(aliyun_rejected, [])
        self.assertEqual(len(oversized_aliyun), 1)
        self.assertEqual(deepl_escaped_batches, [])
        self.assertEqual(len(deepl_escaped_rejected), 1)
        self.assertEqual([len(batch) for batch in volcengine_character_batches], [1, 1])
        self.assertEqual([len(batch) for batch in aliyun_character_batches], [8, 1])

    def test_aliyun_uses_acs_signature_request(self):
        app = build_app(FakeResponse({'Data': {'Translated': 'Hello'}}))
        config = app.provider_configs['aliyun']
        config['access_key_id'] = 'test-id'
        config['access_key_secret'] = 'test-secret'

        result = app._translate_with_provider('aliyun', '你好', 'en-US')

        self.assertTrue(result.success)
        _, _, kwargs = app._http_local.session.requests[0]
        self.assertTrue(kwargs['headers']['Authorization'].startswith('acs test-id:'))
        self.assertEqual(kwargs['headers']['X-Acs-Version'], '2019-01-02')

    def test_aliyun_batch_maps_results_by_returned_index(self):
        app = build_app(FakeResponse({
            'Code': 200,
            'TranslatedList': [
                {'index': '1', 'translated': 'World'},
                {'index': '0', 'translated': 'Hello'},
            ],
        }))
        config = app.provider_configs['aliyun']
        config['access_key_id'] = 'test-id'
        config['access_key_secret'] = 'test-secret'

        results = app._translate_provider_text_batch(
            'aliyun',
            ['你好', '世界'],
            'en-US'
        )

        self.assertEqual([result.text for result in results], ['Hello', 'World'])
        self.assertTrue(all(result.success for result in results))
        _, url, kwargs = app._http_local.session.requests[0]
        self.assertEqual(url, 'https://mt.cn-hangzhou.aliyuncs.com')
        self.assertEqual(kwargs['data']['Action'], 'GetBatchTranslate')
        self.assertIn('Signature', kwargs['data'])
        self.assertEqual(
            json.loads(kwargs['data']['SourceText']),
            {'0': '你好', '1': '世界'}
        )

    def test_failed_results_are_not_cached(self):
        app = build_app(FakeResponse({}))

        def fail(*_args, **_kwargs):
            raise TranslationProviderError('credential rejected')

        app._translate_provider_once = fail
        result = app._translate_with_provider('google', '你好', 'en')

        self.assertFalse(result.success)
        self.assertEqual(app.translation_cache, {})
        self.assertEqual(result.error, 'credential rejected')

    def test_legacy_baidu_credentials_are_migrated(self):
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / 'translation_config.json'
            config_path.write_text(
                json.dumps(
                    {
                        'api_type': 'baidu',
                        'baidu_app_id': 'legacy-app',
                        'baidu_secret_key': 'legacy-secret',
                    }
                ),
                encoding='utf-8'
            )
            app = TranslationApp.__new__(TranslationApp)
            app.default_langs = []
            app.api_type = 'google'
            app.provider_configs = default_provider_configs()
            app.baidu_app_id = ''
            app.baidu_secret_key = ''
            app.logger = None
            app.log = lambda _message: None
            original_config_file = translation_gui.CONFIG_FILE
            translation_gui.CONFIG_FILE = str(config_path)
            try:
                app.load_config()
            finally:
                translation_gui.CONFIG_FILE = original_config_file

            self.assertEqual(app.api_type, 'baidu')
            self.assertEqual(app.provider_configs['baidu']['app_id'], 'legacy-app')
            self.assertEqual(app.provider_configs['baidu']['secret_key'], 'legacy-secret')

    def test_translation_table_save_uses_a_complete_replacement_file(self):
        with tempfile.TemporaryDirectory() as directory:
            app = TranslationApp.__new__(TranslationApp)
            app.table_dir = directory
            app.translation_table = {'测试': {'USA': 'Test'}}
            app.logger = None
            app.log = lambda _message: None

            app.save_translation_table()

            table_path = Path(directory) / 'translation_table.json'
            self.assertEqual(
                json.loads(table_path.read_text(encoding='utf-8')),
                {'测试': {'USA': 'Test'}}
            )
            self.assertEqual(list(Path(directory).glob('.translation_table.*.tmp')), [])


if __name__ == '__main__':
    unittest.main()
