import json
import pytest


def test_plugin_terminal_remains_eager_with_discovery_enabled(tmp_path, monkeypatch):
    monkeypatch.setenv('HERMES_HOME', str(tmp_path))
    plugin = tmp_path / 'plugins' / 'terminal-protocol-test'
    plugin.mkdir(parents=True)
    (tmp_path / 'config.yaml').write_text('plugins:\n  enabled: [terminal-protocol-test]\n')
    (plugin / 'plugin.yaml').write_text('name: terminal-protocol-test\nversion: 0.1.0\ndescription: Test terminal protocol\n')
    (plugin / '__init__.py').write_text('''
def register(ctx):
    ctx.register_tool(name='terminal_test_response', toolset='terminal-protocol-test',
        schema={'name': 'terminal_test_response', 'parameters': {'type': 'object', 'properties': {}}},
        handler=lambda args, **kw: '{"ok":true}', always_visible=True)
    ctx.register_tool(name='ordinary_test_lookup', toolset='terminal-protocol-test',
        schema={'name': 'ordinary_test_lookup', 'parameters': {'type': 'object', 'properties': {}}},
        handler=lambda args, **kw: '{}')
    ctx.register_tool(name='unavailable_test_response', toolset='terminal-protocol-test',
        schema={'name': 'unavailable_test_response', 'parameters': {'type': 'object', 'properties': {}}},
        handler=lambda args, **kw: '{}', check_fn=lambda: False, always_visible=True)
''')
    from hermes_cli.plugins import PluginManager
    from tools.registry import registry
    from tools.tool_search import assemble_tool_defs, ToolSearchConfig, is_deferrable_tool_name
    manager = PluginManager()
    manager.discover_and_load()
    entry = registry.get_entry('terminal_test_response')
    assert entry is not None and entry.always_visible
    assert not is_deferrable_tool_name('terminal_test_response', frozenset({'terminal_test_response'}))
    assert is_deferrable_tool_name('ordinary_test_lookup')
    definitions = registry.get_definitions({'terminal_test_response', 'ordinary_test_lookup', 'unavailable_test_response'})
    result = assemble_tool_defs(definitions, context_length=200000,
                               config=ToolSearchConfig.from_raw({'enabled': 'on'}))
    names = {tool['function']['name'] for tool in result.tool_defs}
    assert result.activated
    assert 'terminal_test_response' in names
    assert 'ordinary_test_lookup' not in names
    assert 'unavailable_test_response' not in names
    assert json.loads(registry.dispatch('terminal_test_response', {}))['ok']
    assert registry.get_definitions(set()) == []
