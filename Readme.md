# Output Formatter Plugin

**Smart formatting system for ISO 26262 work products in Cheshire Cat AI**

Automatically formats raw AI output into professional, ISO-compliant documentation with tables, sections, and workflow guidance.

---

## Features

- ✅ **Automatic Formatting** - Detects workflow stage and applies appropriate formatter
- ✅ **LLM-Enhanced** - Uses LLM to intelligently structure content
- ✅ **ISO 26262 Compliance** - Formats per standard requirements
- ✅ **Multi-Plugin Support** - Works with HARA, FSC, and other safety plugins
- ✅ **Workflow Guidance** - Adds "Next Steps" recommendations
- ✅ **Smart Cleanup** - Prevents duplicate formatting

---

## How It Works

### Automatic Detection

The plugin uses a **hook** that intercepts AI responses before they reach the user:

```python
@hook(priority=5)
def before_cat_sends_message(message, cat):
    # 1. Check working memory for 'needs_formatting' flag
    # 2. Route to appropriate formatter based on 'last_operation'
    # 3. Format content with LLM
    # 4. Add workflow guidance
    # 5. Cleanup state to prevent re-formatting
```

### Workflow Integration

Other plugins signal the formatter by setting working memory flags:

```python
# In HARA plugin after extracting functions:
cat.working_memory['last_operation'] = 'function_extraction'
cat.working_memory['needs_formatting'] = True
```

---

## Supported Formatters

### HARA Operations
| Operation | Formatter | Output |
|-----------|-----------|--------|
| `function_extraction` | `hara_functions` | Function table with IDs |
| `hazop_analysis` | `hara_hazop` | HAZOP matrix (Guide Words × Functions) |
| `operational_situations_defined` | `hara_situations` | Operational scenarios table |
| `esc_assessment_complete` | `hara_esc` | E/S/C ratings matrix |
| `asil_determination_complete` | `hara_asil` | ASIL distribution + table |
| `safety_goals_derived` | `hara_safety_goals` | Safety goals with ASIL |

### FSC Operations
| Operation | Formatter | Output |
|-----------|-----------|--------|
| `strategy_development` | `safety_strategies` | Safety strategies table |
| `fsr_derivation` | `fsrs` | FSR requirements matrix |
| `allocation_complete` | `allocation` | Allocation matrix |
| `mechanisms_identified` | `mechanisms` | Safety mechanisms catalog |
| `validation_criteria_specified` | `validation` | Validation criteria |
| `fsc_verified` | `verification` | Verification report |

---

## Installation

1. Copy plugin folder to Cheshire Cat plugins directory:
   ```
   cat/plugins/AI_Agent-OutputFormatter/
   ```

2. No additional dependencies required (uses Cheshire Cat's LLM)

3. Restart Cheshire Cat

4. Plugin automatically activates when other plugins set formatting flags

---

## File Structure

```
AI_Agent-OutputFormatter/
├── plugin.json               # Plugin metadata
├── README.md                 # This file
└── code/
    └── hooks/
        └── output_formatter.py   # Main hook with formatters
```

---

## Example Transformation

**Before (Raw):**
```
I extracted 5 functions: F-001 Cell Voltage Monitoring, F-002 Temperature Sensing...
```

**After (Formatted):**
```markdown
## 🔧 Safety-Relevant Functions: Battery Management System

| Function ID | Description | Safety Relevance |
|-------------|-------------|------------------|
| F-001 | Cell Voltage Monitoring | Critical - Overcharge protection |
| F-002 | Temperature Sensing | Critical - Thermal runaway prevention |
...

### 🚀 Next Steps
**Recommended:** `apply hazop to functions`
```

---

## Adding New Formatters

### Step 1: Add Operation Mapping

```python
OPERATION_TO_FORMATTER = {
    # ... existing mappings
    'my_new_operation': 'my_formatter',
}
```

### Step 2: Create Formatter Method

```python
def _format_my_content(self, content: str, cat) -> str:
    data = cat.working_memory.get('my_data')
    
    prompt = f"""Format this content professionally.

CONTENT: {content}
DATA: {data}

OUTPUT FORMAT:
## My Formatted Output
...

Only output the formatted result."""
    
    return self.llm(prompt)
```

### Step 3: Add to Router

```python
def format_content(self, content: str, formatter_type: str, cat) -> str:
    if formatter_type == 'my_formatter':
        return self._format_my_content(content, cat)
    # ... existing routes
```

### Step 4: Signal from Plugin

```python
# In your plugin tool:
cat.working_memory['last_operation'] = 'my_new_operation'
cat.working_memory['needs_formatting'] = True
```

---

## Best Practices

✅ **Do:**
- Set `needs_formatting = True` when raw data needs formatting
- Use descriptive operation names
- Include relevant data in working memory
- Clear operation-specific flags after use

❌ **Don't:**
- Set formatting flag if output is already formatted
- Format multiple times for same operation
- Keep stale flags in working memory
- Override cleanup mechanism

---

## Troubleshooting

**Formatter not triggering:**
- Check `needs_formatting` flag is set
- Verify `last_operation` is in `OPERATION_TO_FORMATTER`
- Ensure content length > 50 characters

**Formatting twice:**
- Cleanup not running (check logs)
- Multiple plugins setting flags
- Flag not cleared properly

**Wrong formatter:**
- Check `last_operation` value
- Verify mapping in `OPERATION_TO_FORMATTER`
- Review working memory state

**Debug:**
```python
# Check working memory:
print(cat.working_memory.get('last_operation'))
print(cat.working_memory.get('needs_formatting'))

# Check logs:
# Look for: "🎨 Formatting with [type] formatter"
# Look for: "🧹 Cleared needs_formatting flag"
```

---

## Performance

- **LLM Calls:** 1 per formatted response
- **Hook Priority:** 5 (runs before message sent)
- **Memory Usage:** Minimal (only stores flags)
- **Latency:** +1-2 seconds for LLM formatting

---

## Compatibility

- **Cheshire Cat:** v1.5.0+
- **Python:** 3.9+
- **Works with:**
  - HARA Assistant Plugin
  - FSC Developer Plugin
  - Item Definition Developer Plugin
  - Any plugin that sets working memory flags

---
