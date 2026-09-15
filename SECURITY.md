# 🔒 Security Policy

## 📋 Overview

The **CardioRisk AI** team takes data privacy, model integrity, and software security with utmost seriousness. Because this platform is designed for health telemetry, cardiovascular risk estimation, and clinical decision support, maintaining a secure and reliable codebase is paramount.

---

## 🛡️ Supported Versions

We actively provide security patches and dependency updates for the following versions:

| Version | Supported          | Status |
| ------- | ------------------ | ------ |
| `3.x.x` | :white_check_mark: | Current Active Release |
| `2.x.x` | :x:                | End-of-Life |
| `< 2.0` | :x:                | Unsupported |

---

## 🚨 Reporting a Vulnerability

If you discover a security vulnerability, security flaw, or sensitive data exposure issue within CardioRisk AI, **please do not open a public issue.** 

Instead, follow our responsible disclosure process:

### 1. Contact Information
- **Email**: [tondays52@gmail.com](mailto:tondays52@gmail.com)
- **Subject Line**: `[SECURITY VULNERABILITY] CardioRisk AI - <Brief Description>`

### 2. What to Include in Your Report
To help us triage and resolve the issue quickly, please include:
- A clear summary of the potential vulnerability.
- Step-by-step instructions (or proof-of-concept scripts) to reproduce the issue.
- The affected component (e.g., FastAPI backend, JWT authentication, DSP pipelines, WebSocket telemetry, or Streamlit frontend).
- Any potential impact or attack scenario.
- Suggested mitigations or patches (if available).

### 3. Response Timelines
- **Initial Acknowledgment**: Within **24–48 hours**.
- **Assessment & Triage**: Within **3–5 business days**.
- **Fix & Public Disclosure**: We aim to resolve critical issues within **14 calendar days** before coordinated public disclosure.

---

## 🔐 Security Best Practices for Deployment

When running CardioRisk AI in production or clinic environments:

1. **Environment Credentials & Secrets**:
   - Never commit `.env` files or cryptographic private keys to version control.
   - Use strong, cryptographically generated secrets for `JWT_SECRET_KEY`.

2. **Network & Transport Security**:
   - Always enforce HTTPS / TLS 1.3 for FastAPI REST and WebSocket endpoints.
   - Isolate backend microservices behind a secure reverse proxy (e.g., NGINX, Cloudflare).

3. **PHI & Data Privacy (HIPAA / GDPR Compliance)**:
   - Ensure all patient biometric signals (ECG, PPG, PCG, SCG, and Retinal scans) are encrypted in transit and at rest (AES-256).
   - De-identify patient data prior to utilizing external analytical tools or report generation.

4. **Dependency Auditing**:
   - Regularly audit Python packages via automated tools:
     ```bash
     pip install safety pip-audit
     pip-audit
     ```

---

## 📜 Medical & Regulatory Disclaimer

CardioRisk AI is developed for research, education, and assistive screening purposes. It is not approved as a medical device for standalone autonomous diagnosis. Always adhere to local healthcare regulations, institutional review boards (IRB), and patient consent standards.
