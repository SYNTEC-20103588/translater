import hashlib
import json
import queue
import tempfile
import threading
import time
import unittest
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
    app.abbreviate_translations = lambda: None
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


class TranslationProviderTests(unittest.TestCase):
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
