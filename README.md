# AISheeter Frontend (Google Apps Script)

> Google Apps Script frontend for AISheeter - AI-powered Google Sheets™ add-on

**Version:** 2.3.0  
**Updated:** January 10, 2026

### What's New in v2.3.0

🤖 **AI Agent with Model Selection (Cursor-like)**
- Conversational interface for bulk operations
- **Model selector dropdown** - choose your model like in Cursor IDE
- **Task history** - last 10 tasks persisted, rerun with one click
- Preferences stored in UserProperties (survives sessions)
- Natural language parsing: "Translate A2:A50 to Spanish in B"

🏗️ **Modular HTML Architecture** (v2.2.0)
- Refactored 1,478-line Sidebar.html into clean modules
- Separated concerns: styles, utils, bulk, settings, agent
- Easy to maintain and extend
- Uses GAS `<?!= include() ?>` pattern

🚀 **Best-in-Class Bulk Processing** (v2.1.0)
- Smart auto-detection of selected ranges
- Preview first 3 rows before processing
- Live incremental writes (results appear as they complete)
- Retry failed rows only (not entire job)
- Job persistence (survives sidebar close)

---

## 📁 File Structure

```
ai-sheet-front-end/
│
├── 📜 GOOGLE APPS SCRIPT FILES
├── Code.gs              # Main entry point + AI functions
├── Config.gs            # Environment configuration (LOCAL/PRODUCTION)
├── ApiClient.gs         # HTTP request utilities
├── Crypto.gs            # API key encryption/decryption
├── User.gs              # User identity, settings, status
├── Prompts.gs           # Saved prompts CRUD operations
├── Jobs.gs              # Bulk processing & async jobs
├── Context.gs           # Context engineering & task inference
├── Agent.gs             # AI Agent backend (NEW in v2.3.0)
│
├── 🎨 MODULAR HTML FILES
├── Sidebar.html         # Main orchestrator (~170 lines)
├── Sidebar_Styles.html  # CSS + Tailwind config (~370 lines)
├── Sidebar_Utils.html   # Toast, menu, helpers, home page (~200 lines)
├── Sidebar_Bulk.html    # Bulk Agent feature (~450 lines)
├── Sidebar_Settings.html# Settings, forms, contact (~280 lines)
├── Sidebar_Agent.html   # Conversational AI Agent (NEW in v2.3.0)
├── PromptManager.html   # Prompt management modal
│
└── README.md            # This file
```

### Module Architecture

```
┌─────────────────────────────────────────────────────────┐
│  Sidebar.html (orchestrator)                            │
│  └─ <?!= include('Sidebar_Styles') ?>   → CSS          │
│  └─ <?!= include('Sidebar_Utils') ?>    → Core utils   │
│  └─ <?!= include('Sidebar_Bulk') ?>     → Bulk Agent   │
│  └─ <?!= include('Sidebar_Settings') ?> → Settings     │
│  └─ <?!= include('Sidebar_Agent') ?>    → AI Agent     │
└─────────────────────────────────────────────────────────┘
```

### Storage Architecture

```
┌─────────────────────────────────────────────────────────┐
│  AGENT DATA STORAGE (Lean Architecture)                 │
├─────────────────────────────────────────────────────────┤
│                                                          │
│  SESSION (Memory)          UserProperties      Supabase  │
│  ─────────────────        ──────────────      ────────── │
│  • Current messages        • Agent model       • Jobs    │
│  • Active plan             • Last 10 tasks     • Results │
│  • Temp state              • Preferences       • Usage   │
│                                                          │
│  Lost on refresh ───────► Persists ─────────► Permanent  │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

### Adding New Modules

1. Create `Sidebar_NewFeature.html` with your JavaScript
2. Add `<?!= include('Sidebar_NewFeature') ?>` in `Sidebar.html`
3. Functions become globally available
4. Remember: Utils must be included before features (dependency order)

---

## 🔧 Configuration

### Switching Environments

Edit `Config.gs` line 32:

```javascript
// For local development:
var CURRENT_ENV = ENV_TYPE.LOCAL;

// For production:
var CURRENT_ENV = ENV_TYPE.PRODUCTION;
```

### Environment URLs

| Environment | Base URL |
|-------------|----------|
| **LOCAL** | `http://localhost:3000` |
| **PRODUCTION** | `https://aisheet.vercel.app` |

### Debug Mode

Debug mode enables:
- Request/response logging
- Detailed error messages
- Configuration summary on load

Automatically enabled in LOCAL, disabled in PRODUCTION.

---

## 📦 Module Overview

### Config.gs
Centralized configuration for all environment settings.

```javascript
// Get current environment
Config.getEnv();           // 'LOCAL' or 'PRODUCTION'
Config.isProduction();     // boolean
Config.isLocal();          // boolean

// Get URLs
Config.getBaseUrl();       // 'https://aisheet.vercel.app'
Config.getApiUrl('QUERY'); // 'https://aisheet.vercel.app/api/query'

// Get defaults
Config.getDefaultModel('CHATGPT');  // 'gpt-5-mini'
Config.getSupportedProviders();     // ['CHATGPT', 'CLAUDE', 'GROQ', 'GEMINI']
```

### ApiClient.gs
HTTP request utilities with logging and error handling.

```javascript
// GET request
ApiClient.get('MODELS');
ApiClient.get('PROMPTS', { user_id: '123' });

// POST request
ApiClient.post('QUERY', { model: 'CHATGPT', input: 'Hello' });

// PUT request
ApiClient.put('PROMPTS', { id: '1', name: 'Updated' });

// DELETE request
ApiClient.delete('PROMPTS', { id: '1' });
```

### Crypto.gs
AES encryption for API keys using cCryptoGS library.

```javascript
// Encrypt/decrypt
var encrypted = Crypto.encrypt('sk-...');
var decrypted = Crypto.decrypt(encrypted);

// Salt management
Crypto.hasSalt();        // Check if salt exists
Crypto.regenerateSalt(); // Generate new salt (breaks existing keys!)
```

### User.gs
User identity and settings management.

```javascript
// Identity
getUserEmail();      // Get current user's email
getUserId();         // Get or create user ID from backend

// Settings
getUserSettings();   // Get decrypted settings from backend
saveAllSettings(settings);  // Save encrypted settings to backend

// Credit tracking
logCreditUsage(0.0025);
getCreditUsageLogs();
```

### Prompts.gs
Saved prompts CRUD operations.

```javascript
getSavedPrompts();                      // Get all user's prompts
savePrompt(name, prompt, variables);    // Create new prompt
updatePrompt(id, name, prompt, vars);   // Update existing
deletePrompt(id);                       // Delete prompt
openPromptManager();                    // Open dialog
```

### Jobs.gs
Best-in-class bulk processing & async jobs.

```javascript
// === SMART AUTO-DETECTION ===
var selection = getSelectedRange();
// { range: "A2:A500", rowCount: 498, suggestedOutput: "B", hasData: true }

// === PREVIEW MODE ===
var preview = previewBulkJob("A2:A10", "Summarize: {{input}}", "GEMINI");
// { previews: [{input, output, success}], allSuccessful: true, totalRows: 500 }

// === JOB CREATION ===
var job = createBulkJob(inputData, "Summarize: {{input}}", "GEMINI");

// === STATUS TRACKING ===
var status = getJobStatus(job.id);
// { progress: 45, successCount: 43, errorCount: 2, results: [...], errors: [...] }

// === INCREMENTAL WRITES (live updates) ===
writeIncrementalResults(job.id, results, "B", 2, true);
// Writes only NEW results, highlights green/red

// === ERROR RECOVERY ===
retryFailedRows(job.id);        // Retry only failed rows
var errors = getJobErrors(job.id);  // Get error details

// === JOB PERSISTENCE ===
var active = checkActiveJobs();  // Resume on sidebar reopen
// { hasActiveJob: true, job: {...} }

// === HELPERS ===
var values = getRangeValues("A2:A100");
var count = countNonEmptyRows("A2:A100");
clearResultHighlights("B", 2, 100);  // Remove cell colors
var estimate = estimateBulkCost(100, "GEMINI");
```

### Context.gs
Context engineering with automatic task type inference.

```javascript
// Infer task type from prompt
var taskType = inferTaskType("Extract the email from this text");
// Returns: "EXTRACT"

// Get task info
var info = getTaskTypeInfo("SUMMARIZE");
// { label: "Summarize", icon: "📝", description: "Condense information" }

// Get model recommendations
var recs = getRecommendedModels("CODE");
// { fast: "GROQ", quality: "CLAUDE" }

// Task types: EXTRACT, SUMMARIZE, CLASSIFY, TRANSLATE, CODE, ANALYZE, FORMAT, GENERAL
```

### Agent.gs *(v2.3.0)*
AI Agent backend with preferences, task history, and real-time job updates.

```javascript
// === MODEL SELECTION (Cursor-like) ===
getAgentModel();              // Get current model: 'GEMINI'
setAgentModel('CLAUDE');      // Persists to UserProperties
getAgentModels();             // Get all available models with pricing

// === PREFERENCES ===
getAgentPreferences();        
// { model: 'GEMINI', recentTasks: [{...}] }

// === TASK HISTORY ===
addTaskToHistory(task);       // Add task to history (keeps last 10)
getTaskHistory(5);            // Get last N tasks
clearTaskHistory();           // Clear all history

// === AGENT EXECUTION ===
executeAgentPlan(plan);       // Execute parsed plan, creates jobs
// plan: { inputRange, outputColumns, prompt, model, taskType }

// === REAL-TIME JOB UPDATES (SSE) ===
getJobStreamUrl(['job1', 'job2']); // Get SSE stream URL for job monitoring
// Returns: { url: 'https://.../api/jobs/stream?...', supportsSSE: true, fallbackPollingMs: 2000 }

// === MULTI-LANGUAGE SUPPORT ===
createTranslationPrompts(['Spanish', 'French'], ['B', 'C']);
// Returns column-to-prompt mapping for multi-language translation
```

### Code.gs
Main entry point with core functionality.

```javascript
// Menu/Sidebar
onOpen();       // Initialize menu
showSidebar();  // Display sidebar

// AI Queries (custom functions)
ChatGPT(prompt, imageUrl, model);
Claude(prompt, imageUrl, model);
Groq(prompt, imageUrl, model);
Gemini(prompt, imageUrl, model);

// Image Generation
DALLE(prompt);

// Other
fetchModels();              // Get available models
submitContactForm(name, email, message);
```

---

## 🎨 UI Design

### Design System
- **Font:** Inter (Google Fonts)
- **CSS Framework:** Tailwind CSS via Play CDN
- **Components:** shadcn/ui-inspired

### Key Components
- Cards with subtle shadows
- Button variants (primary, secondary, outline, ghost, destructive)
- Form inputs with focus rings
- Tab navigation
- Toast notifications
- Loading skeletons
- Slide-out menu panel

---

## 🔌 API Endpoints

All endpoints are relative to the base URL.

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/query` | POST | AI text generation |
| `/api/generate-image` | POST | Image generation |
| `/api/models` | GET | Get available models |
| `/api/get-or-create-user` | POST | Create/get user |
| `/api/get-user-settings` | GET | Get user settings |
| `/api/save-all-settings` | POST | Save settings |
| `/api/prompts` | GET/POST/PUT/DELETE | Prompt CRUD |
| `/api/jobs` | GET/POST/DELETE | Bulk jobs |
| `/api/contact` | POST | Contact form |

---

## 📋 Version Tracking

Every file has a standardized header for easy version tracking when copying to Google Apps Script:

```javascript
/**
 * @file Code.gs
 * @version 2.2.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.2.0 (2026-01-10): Added include() for modular HTML
 * - 2.1.0 (2026-01-10): Input validation for empty prompts
 * - 2.0.0 (2026-01-09): Initial modular architecture
 */
```

### Current Versions

| File | Version | Last Updated |
|------|---------|--------------|
| `Code.gs` | 2.2.0 | 2026-01-10 |
| `Config.gs` | 2.3.0 | 2026-01-11 |
| `ApiClient.gs` | 2.0.0 | 2026-01-10 |
| `Crypto.gs` | 2.0.0 | 2026-01-10 |
| `User.gs` | 2.1.0 | 2026-01-10 |
| `Prompts.gs` | 2.0.0 | 2026-01-10 |
| `Jobs.gs` | 2.2.0 | 2026-01-10 |
| `Context.gs` | 2.0.0 | 2026-01-10 |
| `Agent.gs` | 2.3.0 | 2026-01-11 |
| `Sidebar.html` | 2.3.0 | 2026-01-10 |
| `Sidebar_Agent.html` | 2.6.0 | 2026-01-11 |
| `Sidebar_Styles.html` | 2.3.0 | 2026-01-10 |
| `Sidebar_Utils.html` | 2.2.0 | 2026-01-10 |
| `Sidebar_Bulk.html` | 2.2.0 | 2026-01-10 |
| `Sidebar_Settings.html` | 2.2.0 | 2026-01-10 |
| `PromptManager.html` | 2.0.0 | 2026-01-10 |

### When Updating

1. Increment version number (semver: major.minor.patch)
2. Update the `@updated` date
3. Add changelog entry at the top
4. Update the table above

---

## 🚀 Deployment

### Deploy to Google Apps Script

1. Open Google Sheets
2. Extensions → Apps Script
3. Copy all `.gs` and `.html` files (check versions match above)
4. Set `CURRENT_ENV = ENV_TYPE.PRODUCTION` in Config.gs
5. Save and test

### Local Development

1. Start backend locally: `cd ai-sheet-backend && npm run dev`
2. Set `CURRENT_ENV = ENV_TYPE.LOCAL` in Config.gs
3. Test from Google Sheets

---

## 📝 Default Models (January 2026)

| Provider | Default Model | Cost (per MTok) |
|----------|---------------|-----------------|
| **ChatGPT** | gpt-5-mini | $0.25 / $2.00 |
| **Claude** | claude-haiku-4-5 | $1.00 / $5.00 |
| **Groq** | meta-llama/llama-4-scout-17b-16e-instruct | $0.11 / $0.34 |
| **Gemini** | gemini-2.5-flash | $0.075 / $0.30 |

---

## 🔐 Dependencies

### Apps Script Libraries

| Library | ID | Version | Purpose |
|---------|-------|---------|---------|
| cCryptoGS | `1DrAWO6wSBwNgKaIH6hN4njIz2... ` | Latest | AES encryption |

---

## 📄 License

MIT License - See root LICENSE file
