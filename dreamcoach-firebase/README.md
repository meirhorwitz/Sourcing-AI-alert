# DreamCoach Firebase App

This directory contains a simple Firebase setup for the DreamCoach web application.

## Structure

- `public/` - static assets served by Firebase Hosting
- `functions/` - Cloud Functions handling dream submissions

## Setup

1. Install Firebase CLI and initialize functions dependencies:
   ```bash
   npm install -g firebase-tools
   cd dreamcoach-firebase/functions && npm install
   ```
2. Configure required secrets:
   ```bash
   firebase functions:config:set openai.key="YOUR_OPENAI_KEY" smtp.user="YOUR_EMAIL" smtp.pass="EMAIL_PASSWORD"
   ```
3. Deploy:
   ```bash
   firebase deploy
   ```

The frontend submits dreams to the `submitDream` HTTP function which analyzes the dream with OpenAI, stores it in Firestore, and emails the user their analysis.
