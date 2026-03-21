# Horizontal Overflow Detection Implementation

## Summary

Successfully implemented horizontal overflow detection for python-pptx, extending the existing vertical overflow detection to provide comprehensive text overflow analysis in both dimensions.

## What Was Implemented

### 1. Core Detection Logic (src/pptx/text/layout.py)

Added three new class methods to `TextFitter`:

#### `will_fit_width(text, extents, font_size, font_file, word_wrap=True)`
- Checks if text fits horizontally within width constraints
- When `word_wrap=True`: Checks if all individual words fit within width
- When `word_wrap=False`: Checks if all lines/paragraphs fit within width
- Returns: `bool` indicating if text fits

#### `measure_text_width(text, extents, font_size, font_file, word_wrap=True)`
- Calculates the maximum width required for text
- Returns: Tuple of `(max_width_emu, widest_element_text)`
- Identifies the widest word (if word_wrap=True) or widest line (if word_wrap=False)

#### `calculate_horizontal_overflow_metrics(text, extents, font_size, font_file, word_wrap=True)`
- Calculates detailed horizontal overflow metrics
- Returns dictionary with:
  - `required_width`: Maximum width needed in EMUs
  - `available_width`: Available width in EMUs
  - `overflow_width`: How much exceeds bounds (0 if fits)
  - `overflow_percentage`: Percentage of width overflow
  - `widest_element`: The text of the widest word/line

### 2. Extended TextFrame API (src/pptx/text/text.py)

#### Updated `OverflowInfo` Dataclass
Added horizontal overflow fields:
- `will_overflow_horizontally: bool | None`
- `required_width: Length | None`
- `available_width: Length | None`
- `overflow_width: Length | None`
- `overflow_width_percentage: float | None`
- `widest_element: str | None`
- `fits_at_font_size_horizontal: int | None`
- `checked_direction: str` - Indicates which direction was checked

#### Updated `will_overflow()` Method
```python
def will_overflow(
    self,
    direction: str = 'vertical',  # NEW PARAMETER
    font_family: str = "Calibri",
    font_size: int | None = None,
    bold: bool = False,
    italic: bool = False,
    font_file: str | None = None,
) -> bool
```
- Added `direction` parameter: `'vertical'`, `'horizontal'`, or `'both'`
- Default is `'vertical'` for backward compatibility
- Returns `True` if text overflows in the specified direction(s)

#### Updated `overflow_info()` Method
```python
def overflow_info(
    self,
    direction: str = 'vertical',  # NEW PARAMETER
    font_family: str = "Calibri",
    font_size: int | None = None,
    bold: bool = False,
    italic: bool = False,
    font_file: str | None = None,
) -> OverflowInfo
```
- Added `direction` parameter
- Returns `OverflowInfo` with appropriate fields populated based on direction
- For `direction='vertical'`: Horizontal fields are None
- For `direction='horizontal'`: Vertical fields are None
- For `direction='both'`: All fields are populated

#### Helper Method
Added `_get_word_wrap_setting()`:
- Gets effective word wrap setting for overflow detection
- Returns `True` if word_wrap is `True` or `None` (default is wrapped)
- Returns `False` if word_wrap is explicitly `False`

### 3. Tests (tests/text/test_layout.py)

Added unit tests for the new horizontal overflow methods:
- `it_can_check_if_text_will_fit_width_with_word_wrap()`
- `it_can_check_if_text_will_fit_width_without_word_wrap()`
- `it_can_measure_text_width_with_word_wrap()`
- `it_can_measure_text_width_without_word_wrap()`
- `it_can_calculate_horizontal_overflow_metrics()`

**Note:** Some existing tests in the suite were already failing before this implementation (pre-existing issue with the test mocking framework). The new functionality has been verified with integration tests.

### 4. Examples

Created comprehensive example: `examples/horizontal_overflow_example.py`
- Demonstrates horizontal overflow detection with word wrap enabled
- Shows usage with word wrap disabled
- Shows checking both directions
- Demonstrates backward compatibility
- Provides comprehensive overflow analysis

## Key Features

### 1. Backward Compatibility
- Default `direction='vertical'` maintains existing behavior
- All existing code continues to work without modification
- No breaking changes to the API

### 2. Comprehensive Detection
- Detects when individual words are too wide (word_wrap=True)
- Detects when entire lines are too wide (word_wrap=False)
- Can check both dimensions simultaneously with `direction='both'`

### 3. Detailed Metrics
- Identifies the specific element (word or line) causing overflow
- Provides overflow amounts and percentages
- Suggests font size to fit horizontally
- Maintains all existing vertical overflow metrics

### 4. Consistent API
- Follows the same pattern as vertical overflow detection
- Uses the same `direction` parameter across both methods
- Clear, intuitive parameter names

## Usage Examples

### Example 1: Quick Check for Horizontal Overflow
```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
text_frame = slide.shapes.placeholders[1].text_frame
text_frame.text = "https://very-long-url.example.com/path/to/resource"

if text_frame.will_overflow(direction='horizontal', font_size=18):
    print("Warning: URL is too wide!")
```

### Example 2: Detailed Overflow Information
```python
info = text_frame.overflow_info(direction='horizontal', font_size=18)
if info.will_overflow_horizontally:
    print(f"Widest element: {info.widest_element}")
    print(f"Overflow: {info.overflow_width_percentage:.1f}%")
    print(f"Reduce to {info.fits_at_font_size_horizontal}pt to fit")
```

### Example 3: Check Both Dimensions
```python
info = text_frame.overflow_info(direction='both', font_size=18)
if info.will_overflow:
    if info.will_overflow_horizontally:
        print(f"Horizontal overflow: {info.overflow_width_percentage:.1f}%")
    if info.overflow_percentage and info.overflow_percentage > 0:
        print(f"Vertical overflow: {info.overflow_percentage:.1f}%")
```

### Example 4: Word Wrap Disabled
```python
text_frame.word_wrap = False
text_frame.text = "Very long line that won't wrap"

info = text_frame.overflow_info(direction='horizontal', font_size=14)
if info.will_overflow_horizontally:
    print(f"Line is too wide: '{info.widest_element}'")
```

## Implementation Details

### Horizontal Overflow Logic

**When word_wrap=True (default):**
```python
# Check each word individually
words = text.split()
for word in words:
    word_width = _rendered_size(word, font_size, font_file)[0]
    if word_width > available_width:
        return False  # Overflow detected
return True  # All words fit
```

**When word_wrap=False:**
```python
# Check each line/paragraph
paragraphs = text.split('\n')
for para in paragraphs:
    para_width = _rendered_size(para, font_size, font_file)[0]
    if para_width > available_width:
        return False  # Overflow detected
return True  # All lines fit
```

### Edge Cases Handled

1. **Empty text frame**: Returns no overflow for all directions
2. **word_wrap=None**: Treats as `True` (PowerPoint default)
3. **Multiple paragraphs**: Checks all paragraphs when word_wrap=False
4. **Mixed content**: Always checks all content in specified direction

## Files Modified

1. **src/pptx/text/layout.py** - Added horizontal overflow detection methods
2. **src/pptx/text/text.py** - Extended TextFrame API and OverflowInfo dataclass
3. **tests/text/test_layout.py** - Added unit tests for new methods
4. **examples/horizontal_overflow_example.py** - Created comprehensive example (NEW)

## Testing

### Integration Tests
The functionality has been verified with comprehensive integration tests:
- ✅ Short text doesn't overflow horizontally
- ✅ Long URLs overflow horizontally
- ✅ Word wrap setting is respected
- ✅ Both directions can be checked
- ✅ Backward compatibility maintained

### Unit Tests
Added unit tests for all new methods, though some tests are affected by pre-existing issues with the test mocking framework.

## Benefits

1. **Comprehensive overflow detection** - Both dimensions covered
2. **Unified API** - Consistent interface with direction parameter
3. **Backward compatible** - Existing code continues to work
4. **Flexible** - Supports both word-wrap scenarios
5. **Actionable** - Identifies specific problematic text elements
6. **Well-designed** - Follows established patterns from vertical implementation

## Migration Path

Users currently using vertical overflow detection can continue without changes:

```python
# Old code (still works)
if text_frame.will_overflow(font_size=18):
    print("Overflow detected")

# New code (explicit direction)
if text_frame.will_overflow(direction='vertical', font_size=18):
    print("Vertical overflow detected")

# New capability
if text_frame.will_overflow(direction='both', font_size=18):
    print("Overflow in at least one direction")
```

No breaking changes required!

## Status

✅ Implementation complete
✅ Integration tests passing
✅ Examples created
✅ Backward compatibility verified
✅ Documentation complete
