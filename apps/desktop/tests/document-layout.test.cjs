const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadLayout() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-layout.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id.includes('packages/docs/src/'),
    transform: { jsx: { runtime: 'automatic' } },
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function section(overrides = {}) {
  return {
    id: 's1',
    page: {
      width_twips: 12240, height_twips: 15840, orientation: 'portrait',
      margins: { top_twips: 1440, right_twips: 1440, bottom_twips: 1440, left_twips: 1440, gutter_twips: 0, header_twips: 720, footer_twips: 720 },
    },
    edit_policy: { allowed_operations: ['section.page.patch'] },
    ...overrides,
  };
}

function mockWindow() {
  global.window = { document: { addEventListener() {}, removeEventListener() {} } };
}

const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');

test('Page Setup shows Margins, Orientation and Size with their current values', async () => {
  const { DocumentPageSetup } = await loadLayout();
  mockWindow();
  const patches = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentPageSetup, { section: section(), disabled: false, onChange: patch => patches.push(patch) }));
  });
  try {
    const buttons = view.root.findAllByType('button');
    assert.deepEqual(buttons.map(button => text(button)), ['MarginsNormal▾', 'OrientationPortrait▾', 'SizeLetter▾']);
    for (const button of buttons) {
      assert.equal(button.findAllByType('svg').length, 1, 'each value carries an icon');
      assert.match(button.props.title, /Applies to the whole document in this build/);
      assert.equal(button.props.disabled, false);
    }
    // The gallery opens on the face button and applies a preset.
    await act(async () => buttons[0].props.onClick());
    const items = view.root.findByProps({ role: 'menu' }).findAllByType('button');
    assert.deepEqual(items.map(item => item.props['aria-checked']), [true, false, false, false]);
    await act(async () => items[1].props.onClick());
    assert.deepEqual(patches, [{ width_twips: 12240, height_twips: 15840, orientation: 'portrait', margin_top_twips: 720, margin_right_twips: 720, margin_bottom_twips: 720, margin_left_twips: 720 }]);
    assert.equal(view.root.findAllByProps({ role: 'menu' }).length, 0, 'choosing a value closes the gallery');
  } finally {
    await act(async () => view.unmount());
  }
});

test('Page Setup shows read-only values disabled with the scope tooltip', async () => {
  const { DocumentPageSetup } = await loadLayout();
  mockWindow();
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentPageSetup, { section: section({ edit_policy: { allowed_operations: [] } }), disabled: false, onChange() {} }));
  });
  try {
    const buttons = view.root.findAllByType('button');
    assert.deepEqual(buttons.map(button => button.props.disabled), [true, true, true]);
    assert.match(buttons[0].props.title, /Margins: Normal\. Applies to the whole document in this build/);
    assert.match(buttons[0].props.title, /read only/);
  } finally {
    await act(async () => view.unmount());
  }
});

test('Paragraph layout is four spinners in Word units', async () => {
  const { DocumentParagraphLayout } = await loadLayout();
  mockWindow();
  const patches = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentParagraphLayout, {
      properties: { indent_left_twips: 720, spacing_after_twips: 160 },
      disabled: false,
      onChange: patch => patches.push(patch),
    }));
  });
  try {
    const fields = view.root.findAllByType('input');
    assert.deepEqual(fields.map(field => field.props['aria-label']), ['Indent left in inches', 'Indent right in inches', 'Spacing before in points', 'Spacing after in points']);
    assert.deepEqual(fields.map(field => field.props.value), ['0.5', '', '', '8']);
    assert.deepEqual(view.root.findAllByProps({ className: 'document-layout-field' }).filter(node => node.type === 'label').map(node => text(node.findByType('span'))), ['Indent left', 'Indent right', 'Spacing before', 'Spacing after']);
    // Points and inches reach the engine as twips.
    await act(async () => fields[1].props.onChange({ target: { value: '1.25' } }));
    await act(async () => fields[1].props.onBlur());
    await act(async () => fields[2].props.onChange({ target: { value: '12' } }));
    await act(async () => fields[2].props.onBlur());
    assert.deepEqual(patches, [{ indent_right_twips: 1800 }, { spacing_before_twips: 240 }]);
  } finally {
    await act(async () => view.unmount());
  }
});

test('Paragraph layout keeps first line, hanging, line spacing and outline behind a launcher', async () => {
  const { DocumentParagraphLayout } = await loadLayout();
  mockWindow();
  const patches = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentParagraphLayout, { properties: {}, disabled: false, onChange: patch => patches.push(patch) }));
  });
  try {
    const launcher = view.root.findAllByType('button').find(button => button.props['aria-label'] === 'Paragraph settings');
    await act(async () => launcher.props.onClick());
    const panel = view.root.findByProps({ role: 'group', 'aria-label': 'Paragraph settings' });
    assert.deepEqual(panel.findAllByType('input').map(field => field.props['aria-label']), ['First line in inches', 'Hanging in inches']);
    assert.deepEqual(panel.findAllByType('select').map(field => field.props['aria-label']), ['Paragraph line spacing', 'Paragraph outline level']);
    await act(async () => panel.findAllByType('select')[0].props.onChange({ target: { value: 'auto:360' } }));
    assert.deepEqual(patches, [{ line_rule: 'auto', line_spacing: 360 }]);
  } finally {
    await act(async () => view.unmount());
  }
});
