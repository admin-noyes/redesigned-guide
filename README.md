# Email Template Builder

A powerful, client-side web application for creating responsive email templates through a visual drag-and-drop interface. No server required to use locally—components are built with vanilla JavaScript and exported as production-ready HTML.

![Template Builder Interface](./assets/images/favicon_io/favicon-32x32.png)

## Table of Contents

- [Features](#features)
- [Getting Started](#getting-started)
- [Architecture](#architecture)
- [Components](#components)
- [Usage Guide](#usage-guide)
- [Property Editor](#property-editor)
- [Exporting Templates](#exporting-templates)
- [File Structure](#file-structure)
- [Technical Details](#technical-details)
- [API Integration](#api-integration)
- [Development](#development)

---

## Features

### Core Template Building
- **Visual Drag & Drop Interface** – Drag components from the left sidebar onto the canvas to build emails
- **18+ Pre-built Components** – Header, Hero Banner, Text Blocks, Buttons, Images, Social Links, Data Tables, and more
- **Live Property Editor** – Right-side panel for real-time customization of component properties
- **Component Reordering** – Drag components on the canvas to rearrange their order

### Editing Capabilities
- **Customizable Properties** – Edit content, colors, alignment, sizes, URLs, and styling for each component
- **Color Picker** – Dual input (color picker + hex text) for precise color selection
- **Content Templates** – Pre-populated default values for quick starting points

### History & Workflow
- **Undo/Redo** – Full undo/redo history (50-level depth)
- **Clear Canvas** – Quickly reset to start a new template
- **Persistent Templates** – Save and load previously created templates (API-backed)
- **Template Naming** – Give templates memorable names when saving

### Export Options
- **Copy HTML** – Copy final template HTML directly to clipboard
- **Download File** – Download template as a `.html` file ready for email clients
- **Email-Client Optimized** – Generated HTML includes Outlook compatibility and responsive styles

### UI/UX Features
- **Dark Mode** – Toggle dark mode for comfortable editing
- **Toast Notifications** – User feedback for all actions (copy, delete, save, etc.)
- **Responsive Layout** – Three-panel interface: components palette, canvas, properties panel
- **Navigation Menu** – Quick access to related pages (CosmosDB, Logic, Repo, Dev QA)

---

## Getting Started

### Installation

1. **Clone or download** the repository
2. **No build step required** – the application runs entirely in the browser
3. **Open** `index.html` in any modern web browser

```bash
# Simply open in a browser (no server needed for basic functionality)
open index.html

# Or if you want a local server:
python -m http.server 8000
# Then visit: http://localhost:8000
```

### Basic Workflow

1. **Add Components** – Drag a component from the left sidebar onto the canvas
2. **Customize** – Click a component to select it and edit properties on the right
3. **Arrange** – Drag components on the canvas to reorder
4. **Preview** – See real-time changes in the canvas
5. **Export** – Copy HTML or download when ready

---

## Architecture

### Technology Stack
- **Vanilla JavaScript** – No frameworks or external dependencies
- **HTML5** – Semantic structure
- **CSS3** – Modern styling with CSS variables for theming
- **Local Storage** – Browser-based state management (when using API)

### Design Pattern
- **Modular Component System** – Each email component has type, properties, and unique ID
- **State Management** – Global component array with undo/redo history
- **Event-Driven UI** – DOM events drive all interactions
- **Functional Rendering** – Components re-render on state changes

### Key Architecture Files

```
├── index.html              # Main application markup
├── boilerplate.html        # Email template boilerplate (reference)
├── assets/
│   ├── js/
│   │   └── script.js       # Core application logic (1470 lines)
│   ├── css/
│   │   ├── application.css # UI styles and layout
│   │   ├── mmit-styles.css # MMIT styleguide overrides
│   │   └── styles.css      # Base styles
│   ├── images/
│   │   ├── favicon_io/     # Favicon assets
│   │   └── mmit_logo.png   # Company logo
│   └── pages/              # Secondary pages
```

---

## Components

The builder includes 18 component types, each with customizable properties:

| Component | Purpose | Key Properties |
|-----------|---------|-----------------|
| **Header** | Email header with logo and company name | logoUrl, companyName, backgroundColor, textColor |
| **Hero Banner** | Large featured image with overlay text | imageUrl, headline, subtext, overlayOpacity |
| **Text Block** | Paragraph content | content, fontSize, textAlign, textColor |
| **Button** | Call-to-action button | text, url, backgroundColor, textColor, borderRadius, align |
| **Image** | Standalone image | url, alt, width (full/75%/50%), align |
| **Divider** | Horizontal line separator | color, thickness, style (solid/dashed/dotted) |
| **Spacer** | Vertical spacing | height (8-120px) |
| **Social Links** | Social media icon links | platforms, iconSize, align |
| **Two Columns** | Side-by-side content layout | leftContent, rightContent, gap |
| **News Header** | Logo + news column (Mailjet style) | logoUrl, newsText, bgColor |
| **CTA Button** | Prominent call-to-action | text, bgColor, textColor |
| **Features (Two)** | Two-column feature block | leftTitle, leftText, rightTitle, rightText |
| **Image Block** | Centered image with sizing | src, alt, width |
| **Data Table** | Tabular data | headers, rows |
| **Title + Paragraphs** | Heading with multiple paragraphs | title, paragraphs |
| **Features (Three)** | Three-column icon features | items (array of {icon, title, text}) |
| **Footer** | Email footer with links | companyName, address, unsubscribeText, unsubscribeUrl, colors |

### Adding Custom Components

To add a new component type:

1. Add default properties in `defaultProps` object (lines ~20-100)
2. Add render case in `renderComponentPreview()` function (~450-800)
3. Add property editor UI in `renderPropertyEditor()` switch statement (~1000-1300)
4. Add component item to the palette in `index.html` (~100-200)

---

## Usage Guide

### Selecting & Editing Components

**Select a component:**
- Click anywhere on a component in the canvas
- The component border highlights in blue
- Properties panel updates on the right

**Edit properties:**
- Modify text in text inputs
- Use color pickers for colors (shows hex value alongside)
- Use range sliders for numeric values
- Changes apply in real-time

**Delete a component:**
- Click the trash icon on the component while it's selected
- Or click the delete button that appears when hovering

### Dragging & Reordering

**Add from palette:**
- Click and drag a component from the left sidebar
- Drop it anywhere on the canvas

**Reorder on canvas:**
- Click the drag handle (≡ icon) on a component
- Drag it to a new position
- Release to reorder

### Undo/Redo

- **Undo** – Click the undo arrow or press (when enabled)
- **Redo** – Click the redo arrow (when enabled)
- History maintains up to 50 states

### Canvas Management

- **Clear Canvas** – Removes all components and resets state
- **Count Display** – Shows number of components in header

---

## Property Editor

The property panel on the right adapts based on selected component:

### Button Properties Example
```
Button Text       [input field]
URL              [input field]
Alignment        [select: Left, Center, Right]
Background Color [color picker + hex input]
Text Color       [color picker + hex input]
Border Radius    [0-24px slider]
```

### Special Property Types

**Text Inputs** – For short strings (text, urls, names)

**Textareas** – For longer content (paragraphs, addresses, descriptions)

**Select Dropdowns** – For predefined options (alignment, style, width)

**Color Pickers** – Dual input (visual picker + hex text synchronized)

**Range Sliders** – For numeric values with limits (size, opacity, spacing)

**Checkboxes** – For toggles (social platform enablement)

### Property Parsing

Complex properties auto-parse when edited:
- **CSV Headers** – `"Col1, Col2, Col3"` → array
- **CSV Rows** – Line-separated CSV → 2D array
- **Line-separated Lists** – Newline delimited → array
- **JSON** – Attempts to parse for array properties

---

## Exporting Templates

### Copy HTML

```javascript
Click "Copy HTML" button → Full template code copied to clipboard
```

### Download File

```javascript
Click "Export" → "Download HTML" → Saves as email-template.html
```

### Export Format

Generated HTML includes:
- Email client compatibility (Outlook MSO tags)
- CSS reset (Meyer reset)
- Responsive meta tags
- Proper table structure for email clients
- Dark mode awareness (if enabled)

### Generated HTML Structure

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Template</title>
  <style>
    /* CSS Reset & Email-safe styles */
  </style>
</head>
<body>
  <table role="presentation">
    <tr>
      <td>
        <!-- Components rendered here -->
      </td>
    </tr>
  </table>
</body>
</html>
```

---

## File Structure

```
template-composer-gsd/
├── index.html                           # Main application (326 lines)
├── boilerplate.html                     # Email template reference (496 lines)
├── README.md                            # This file
├── .vscode/                             # VS Code settings
└── assets/
    ├── css/
    │   ├── application.css              # Main UI styles (809 lines)
    │   ├── mmit-styles.css              # Design system overrides
    │   ├── styles.css                   # Base styles
    │   └── fonts/
    │       └── stylesheet.css           # Custom font definitions
    ├── js/
    │   └── script.js                    # Core logic (1470 lines)
    ├── images/
    │   ├── favicon_io/
    │   │   ├── apple-touch-icon.png
    │   │   ├── favicon-16x16.png
    │   │   ├── favicon-32x32.png
    │   │   └── site.webmanifest
    │   └── mmit_logo.png
    └── pages/
        ├── cosmosDB/                    # CosmosDB info page
        │   ├── index.html
        │   └── mmit-application-styles.css
        ├── dev-qa/                      # Dev/QA testing page
        │   ├── index.html
        │   ├── script.js
        │   └── styles.css
        ├── logic/                       # Logic documentation page
        │   └── index.html
        └── repo/                        # GitHub PR GUI helper
            ├── app.js
            ├── help.html
            ├── index.html
            ├── README.md
            └── styles.css
```

---

## Technical Details

### State Management

```javascript
let components = [];        // Array of component objects
let selectedId = null;      // Currently selected component ID
let darkMode = false;       // Dark mode toggle state
let history = [[]];         // Undo/redo history stack
let historyIndex = 0;       // Current position in history
```

### Component Object Structure

```javascript
{
  id: "comp_abc123xyz",     // Unique identifier
  type: "text",             // Component type
  props: {                  // Type-specific properties
    content: "...",
    fontSize: 16,
    textColor: "#4a4a4a"
  }
}
```

### Key Functions

| Function | Purpose |
|----------|---------|
| `renderCanvas()` | Render all components to the canvas |
| `renderComponentPreview(component)` | Generate HTML for a single component |
| `renderPropertyEditor()` | Render the right-side property panel |
| `generateHtml()` | Generate final email template HTML |
| `selectComponent(id)` | Select a component and update UI |
| `deleteComponent(id)` | Remove component from canvas |
| `pushHistory(components)` | Add state to undo/redo history |

### Drag & Drop Implementation

The application uses native HTML5 drag-and-drop API:
- **Palette Drag** – Components dragged from left sidebar
- **Canvas Drop** – Components added to canvas
- **Canvas Reordering** – Components dragged within canvas to reorder
- **Data Transfer** – Component types passed via `dataTransfer` object

---

## API Integration

The application includes optional API backend support for saving/loading templates:

### API Endpoints

```javascript
GET    /api/templates              // Fetch all templates
POST   /api/templates              // Create new template
PATCH  /api/templates/{id}         // Update template
DELETE /api/templates/{id}         // Delete template
```

### Request/Response Format

**Create/Update Template:**
```json
{
  "name": "Newsletter Q1 2024",
  "components": [
    { "id": "comp_123", "type": "header", "props": {...} },
    { "id": "comp_456", "type": "text", "props": {...} }
  ]
}
```

**Template List Response:**
```json
[
  {
    "id": "template_123",
    "name": "Newsletter Q1 2024",
    "components": [...],
    "createdAt": "2024-01-15T10:30:00Z"
  }
]
```

### Save/Load Workflow

```javascript
// API functions available but require backend:
fetchTemplates()       // GET all templates
saveTemplate(name, components)    // POST new template
updateTemplate(id, data)  // PATCH existing template
deleteTemplate(id)     // DELETE template
```

**Note:** These functions attempt to make fetch requests but gracefully error if no backend is available. The application works fine without API for local use.

---

## Development

### Running Locally

**Option 1: Direct File Open**
```bash
# Just open index.html in your browser
open index.html
```

**Option 2: Local HTTP Server**
```bash
# Python 3
python -m http.server 8000
# Then visit: http://localhost:8000

# Python 2
python -m SimpleHTTPServer 8000

# Node.js
npx http-server
```

### Code Organization

**index.html** (326 lines)
- Semantic HTML structure
- Component palette markup
- Modal dialogs (export, save, load)
- CSS/JS imports

**assets/js/script.js** (1,470 lines)
- Module pattern IIFE for encapsulation
- State management
- Event handling
- Component rendering
- Export logic

**assets/css/application.css** (809 lines)
- CSS custom properties (variables) for theming
- Flexbox layout
- Component styling
- Modal styles
- Toast notification styles

### Key Code Patterns

**Module Pattern (Self-Invoking Function)**
```javascript
(function() {
  'use strict';
  
  // Private state and functions
  let components = [];
  
  function renderCanvas() { ... }
  
  // Initialization
  function init() { ... }
  
  // Auto-run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
```

**Event Delegation**
```javascript
// Edit handlers attach to dynamically created elements
editor.querySelectorAll('input, textarea, select').forEach(el => {
  el.addEventListener('input', (e) => { ... });
});
```

**Immutable State Updates**
```javascript
components = components.map(c => 
  c.id === selectedId ? { ...c, props: newProps } : c
);
```

### Browser Support

- Modern browsers (Chrome, Firefox, Safari, Edge)
- Requires ES6 support (template literals, arrow functions, const/let)
- Drag & Drop API support
- FileReader API (for downloads)
- Clipboard API (for copy)

### Performance Considerations

- Components limit: ~100 before noticeable slowdown
- History depth: 50 states (prevents memory bloat)
- Re-renders on every state change (acceptable for most use cases)
- No virtualization needed for typical email template sizes

---

## Secondary Pages

### Repo Page (`/assets/pages/repo/`)
*GitHub PR GUI Helper*
- Non-developer friendly interface for GitHub operations
- Create branches, commit files, open pull requests
- Uses GitHub API with Personal Access Token
- Requires: `repo` scope PAT
- Files: `index.html`, `app.js`, `styles.css`, `help.html`

### CosmosDB Page (`/assets/pages/cosmosDB/`)
*Database Integration Interface*
- Visual interface for CosmosDB operations
- Component styling shared with main app

### Dev/QA Page (`/assets/pages/dev-qa/`)
*Testing & Validation Tools*
- Development and QA utilities
- Testing interface and validation tools

### Logic Page (`/assets/pages/logic/`)
*Documentation & Logic Reference*
- Application logic documentation
- Reference guide for component behavior

---

## Customization Guide

### Adding a New Component Type

1. **Define Defaults** (script.js, line ~20)
   ```javascript
   const defaultProps = {
     // ... existing components
     'my-component': {
       title: 'Default Title',
       content: 'Default content',
       color: '#ffffff'
     }
   };
   ```

2. **Add to Palette** (index.html, line ~150)
   ```html
   <div class="component-item" draggable="true" data-type="my-component">
     <div class="component-icon"><!-- SVG icon --></div>
     <div class="component-info">
       <span class="component-name">My Component</span>
       <span class="component-desc">Description</span>
     </div>
   </div>
   ```

3. **Add Preview Renderer** (script.js, renderComponentPreview, line ~450)
   ```javascript
   case 'my-component':
     content = `
       <div style="...">
         <h3>${props.title}</h3>
         <p>${props.content}</p>
       </div>
     `;
     break;
   ```

4. **Add Property Editor** (script.js, renderPropertyEditor, line ~1000)
   ```javascript
   case 'my-component':
     html = `
       <div class="property-group">
         <label>Title</label>
         <input type="text" class="input" data-prop="title" value="${props.title}">
       </div>
       <!-- more properties -->
     `;
     break;
   ```

### Theming & Styling

CSS Variables in `application.css`:
```css
:root {
  --app-bg: #f8fafc;           /* Background color */
  --app-bg-card: #ffffff;      /* Card background */
  --app-border: #e2e8f0;       /* Borders */
  --app-text: #0f172a;         /* Text color */
  --app-text-muted: #64748b;   /* Muted text */
}
```

Modify colors here to theme the entire application.

---

## Troubleshooting

### Components not dragging
- Ensure JavaScript is enabled
- Check browser console for errors
- Try using a recent version of Chrome, Firefox, or Safari

### Export not working
- Verify components are present on canvas
- Check browser console for fetch errors
- For API features, ensure backend is running

### Dark mode not showing
- Dark mode toggle is hidden by default (commented out in HTML)
- To enable: Remove `visibility: hidden` from `#btnDarkMode`

### Save/Load not working
- Backend API endpoints required for save/load
- See API Integration section above
- Local usage works without backend

---

## Future Enhancements

Potential features for v2.0:
- Component templates/presets
- CSS custom style editor
- Responsive breakpoint editor
- A/B test variant support
- Email preview in different clients
- Batch component operations
- Component library management
- Collaboration features
- Version history with diffs

---

## License

[Add your license here]

## Support

For issues, feature requests, or contributions:
1. Check the troubleshooting section above
2. Review code comments in `script.js`
3. Inspect browser console for errors
4. Verify file structure matches documentation

---

## Credits

Built with vanilla JavaScript, HTML5, and CSS3. Includes email client compatibility helpers and responsive design patterns optimized for email delivery.

**MMIT Styleguide** integrated for consistent branding and visual design.

---

**Last Updated:** February 2026  
**Version:** 1.0.0