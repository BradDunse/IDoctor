# The I Doctor

Rewrites first-person copy into reader-focused "you" based language using a smart crossfade approach.

---

## How to deploy in plain English

### Step 1 — Put the files on GitHub

1. Go to github.com and sign in (or create a free account).
2. Click the green "New" button to create a new repository.
3. Name it "i-doctor" and click "Create repository".
4. Upload all these files into that repository. You can drag and drop them in the browser.

### Step 2 — Deploy to Vercel

1. Go to vercel.com and sign in with your GitHub account.
2. Click "Add New Project".
3. Find your "i-doctor" repository and click "Import".
4. Before clicking Deploy, click "Environment Variables".
5. Add one variable:
   - Name: ANTHROPIC_API_KEY
   - Value: paste your API key here (the one you created at console.anthropic.com)
6. Click "Deploy".
7. Vercel gives you a live URL. That's your tool on the internet.

### Step 3 — Test it

Visit your new URL, enter an email, paste some text, adjust the sliders, and click "Download my rewrite". A Word document will download with your rewritten copy.

### Step 4 — Connect your email list (optional)

To capture emails into ConvertKit or Mailchimp, open api/rewrite.js and look for the comment that says "EMAIL INTEGRATION GOES HERE". Your developer or Claude can wire that up in a few minutes.

---

## What the sliders do

**Opening voice density (1-10):** Controls how much of the writer's own voice and self-references appear in the first eighth of the piece. Higher means more personal.

**Crossfade transition speed (1-10):** Controls how quickly the writing shifts from personal to "you" based through the middle section. Higher means faster.

Beyond the three-eighths mark the piece defaults to "you" based language throughout.

---

## Costs

The tool uses Claude Sonnet via the Anthropic API. A typical newsletter rewrite costs roughly $0.05 to $0.10. For 500 rewrites a month expect around $25 to $50 in API costs.
