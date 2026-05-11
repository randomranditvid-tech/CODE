import os
import argparse
import pandas as pd

# ---------------------------------------------------------------------------
# Command Line Arguments
# ---------------------------------------------------------------------------

parser = argparse.ArgumentParser(
    description="Update Abstract and Impact based on Vulnerability Title"
)

parser.add_argument(
    "--input",
    required=True,
    help="Input Excel file (.xlsx)"
)

parser.add_argument(
    "--output",
    required=True,
    help="Output Excel file (.xlsx)"
)

args = parser.parse_args()

INPUT_FILE = args.input
OUTPUT_FILE = args.output

# ---------------------------------------------------------------------------
# Vulnerability Mapping
# ---------------------------------------------------------------------------

VULN_DETAILS = {

    "client dom code injection": {
        "abstract": "The application processes untrusted client-side input and passes it to unsafe JavaScript execution sinks such as eval(), innerHTML, or document.write(), enabling arbitrary script execution within the browser context.",
        "impact": "Successful exploitation may allow attackers to execute malicious JavaScript, steal session tokens, manipulate DOM content, perform phishing attacks, or compromise user accounts."
    },

    "client dom stored xss": {
        "abstract": "The application stores untrusted data in client-side storage mechanisms and later renders it in the DOM without proper sanitization or encoding.",
        "impact": "Attackers may execute persistent malicious scripts in users’ browsers, leading to session hijacking, credential theft, defacement, or unauthorized actions."
    },

    "client dom xss": {
        "abstract": "The application writes untrusted client-side data directly into dangerous DOM sinks without proper validation or output encoding.",
        "impact": "Exploitation may enable attackers to execute arbitrary JavaScript in the victim’s browser and compromise user sessions or sensitive information."
    },

    "client hardcoded domain": {
        "abstract": "The application contains hardcoded domain names or endpoint references within client-side code.",
        "impact": "Exposure of internal or environment-specific domains may reveal infrastructure details and increase the risk of targeted attacks or environment misconfiguration."
    },

    "client jquery deprecated symbols": {
        "abstract": "The application uses deprecated jQuery APIs or unsupported library versions that may contain known security weaknesses.",
        "impact": "Attackers may exploit publicly known vulnerabilities in outdated libraries to execute malicious code or bypass security controls."
    },

    "client privacy violation": {
        "abstract": "Sensitive or personally identifiable information (PII) is improperly stored or exposed in client-side components or storage mechanisms.",
        "impact": "Exposure of sensitive data may lead to privacy breaches, regulatory non-compliance, identity theft, or unauthorized disclosure of user information."
    },

    "client use of iframe without sandbox": {
        "abstract": "The application embeds external or untrusted content using iframe elements without applying sandbox restrictions.",
        "impact": "Attackers may exploit unsandboxed iframes to perform clickjacking, malicious redirects, or execute unauthorized actions within the user context."
    },

    "heap inspection": {
        "abstract": "Sensitive information is retained in application memory for extended periods and may be recoverable through memory inspection techniques.",
        "impact": "Attackers with memory access may extract credentials, cryptographic keys, or sensitive business information from application memory."
    },

    "httponlycookies": {
        "abstract": "Session or authentication cookies are not configured with the HttpOnly attribute.",
        "impact": "Malicious client-side scripts may access sensitive cookies, potentially resulting in session hijacking and unauthorized account access."
    },

    "unprotected cookie": {
        "abstract": "Sensitive cookies lack appropriate security attributes such as Secure, HttpOnly, or SameSite protections.",
        "impact": "Attackers may intercept, manipulate, or misuse cookies to compromise user sessions or perform cross-site request forgery attacks."
    },

    "ldap injection": {
        "abstract": "The application constructs LDAP queries using untrusted user input without proper sanitization or escaping.",
        "impact": "Attackers may manipulate LDAP queries to bypass authentication, retrieve unauthorized directory information, or escalate privileges."
    },

    "stored ldap injection": {
        "abstract": "Stored application data is later used in LDAP queries without sufficient validation or escaping.",
        "impact": "Successful exploitation may allow unauthorized directory access, authentication bypass, or exposure of sensitive LDAP information."
    },

    "log forging": {
        "abstract": "The application logs untrusted input without sanitization, allowing attackers to manipulate log entries.",
        "impact": "Attackers may forge or tamper with logs, conceal malicious activities, or mislead incident investigations and auditing processes."
    },

    "path traversal": {
        "abstract": "The application improperly validates file path input, allowing attackers to access files and directories outside the intended location.",
        "impact": "Successful exploitation may expose sensitive files, configuration data, or system resources and potentially lead to further compromise."
    },

    "missing hsts header": {
        "abstract": "The application does not enforce HTTP Strict Transport Security (HSTS) headers for secure HTTPS communication.",
        "impact": "Attackers may exploit protocol downgrade attacks or intercept unencrypted traffic, leading to session compromise or data exposure."
    },

    "no request validation": {
        "abstract": "The application does not properly validate incoming user requests for malicious or unsafe content.",
        "impact": "Lack of request validation may increase the risk of injection attacks, XSS, or malicious payload processing."
    },

    "open redirect": {
        "abstract": "The application redirects users to URLs based on untrusted input without sufficient validation.",
        "impact": "Attackers may exploit the vulnerability for phishing attacks, malicious redirects, or bypassing security controls."
    },

    "password in comment": {
        "abstract": "Sensitive credentials or passwords are exposed within source code comments.",
        "impact": "Exposure of credentials may allow unauthorized access to systems, applications, or sensitive resources."
    },

    "password in configuration file": {
        "abstract": "Passwords or secrets are stored in plaintext within application configuration files.",
        "impact": "Attackers gaining access to configuration files may obtain sensitive credentials and compromise connected systems or databases."
    },

    "use of hardcoded password": {
        "abstract": "Hardcoded credentials are embedded directly within the application source code.",
        "impact": "Exposure of hardcoded secrets may lead to unauthorized access, privilege escalation, or compromise of associated systems and services."
    },

    "privacy violation": {
        "abstract": "Sensitive or personally identifiable information is improperly collected, stored, or exposed by the application.",
        "impact": "Data exposure may result in regulatory violations, reputational damage, financial penalties, or identity theft."
    },

    "prototype pollution": {
        "abstract": "The application improperly handles user-controlled object properties, allowing modification of JavaScript object prototypes.",
        "impact": "Attackers may manipulate application logic, bypass security checks, or execute arbitrary code within the application context."
    },

    "reflected xss all clients": {
        "abstract": "The application reflects untrusted input in server responses without proper encoding or sanitization.",
        "impact": "Attackers may execute malicious scripts in victims’ browsers, leading to session theft, phishing, or unauthorized actions."
    },

    "stored xss": {
        "abstract": "The application stores malicious user-supplied input and later renders it without proper output encoding.",
        "impact": "Stored XSS may affect multiple users and result in persistent client-side attacks, session hijacking, or credential theft."
    },

    "ssl verification bypass": {
        "abstract": "The application disables or bypasses SSL/TLS certificate validation during secure communications.",
        "impact": "Attackers may perform man-in-the-middle attacks, intercept encrypted traffic, or impersonate trusted services."
    },

    "ssrf": {
        "abstract": "The application performs server-side requests using user-controlled input without sufficient validation or restrictions.",
        "impact": "Attackers may access internal systems, cloud metadata services, or restricted network resources through the vulnerable server."
    },

    "trust boundary violation in session variables": {
        "abstract": "The application stores untrusted user input directly into trusted session variables without proper validation.",
        "impact": "Attackers may manipulate trusted application state, bypass security controls, or perform unauthorized actions."
    },

    "unsafe use of target blank": {
        "abstract": "Links using target=\"_blank\" are implemented without rel=\"noopener noreferrer\" protections.",
        "impact": "Attackers may exploit reverse tabnabbing techniques to redirect users to malicious pages or manipulate the originating window."
    },

    "use of broken or risky cryptographic algorithm": {
        "abstract": "The application relies on weak, deprecated, or insecure cryptographic algorithms for security-sensitive operations.",
        "impact": "Weak cryptography may allow attackers to decrypt sensitive data, forge signatures, or compromise authentication mechanisms."
    },

    "csrf": {
        "abstract": "The application does not adequately verify that state-changing requests originate from legitimate authenticated users.",
        "impact": "Attackers may trick authenticated users into performing unauthorized actions such as changing settings, initiating transactions, or modifying data."
    },

    "deserialization of untrusted data": {
        "abstract": "The application deserializes untrusted or user-controlled data without sufficient validation or integrity checks.",
        "impact": "Successful exploitation may lead to remote code execution, privilege escalation, or application compromise."
    },

    "hardcoded password in connection string": {
        "abstract": "Database connection strings contain hardcoded credentials within source code or configuration files.",
        "impact": "Exposure of database credentials may allow attackers to gain unauthorized access to backend databases and sensitive data."
    },

    "sql injection": {
        "abstract": "The application constructs SQL queries using untrusted user input without sufficient sanitization or parameterization.",
        "impact": "Attackers may manipulate database queries to access, modify, or delete sensitive information and potentially compromise the database server."
    },

    "second order sql injection": {
        "abstract": "The application stores untrusted input that is later used in SQL queries without proper sanitization.",
        "impact": "Attackers may exploit stored malicious input to manipulate backend database queries and compromise sensitive data or database integrity."
    },

    "use of cryptographically weak prng": {
        "abstract": "The application uses predictable or non-cryptographically secure random number generators for security-sensitive operations.",
        "impact": "Attackers may predict generated values such as session tokens, reset links, or cryptographic nonces, leading to account compromise."
    },

    "client dom stored code injection": {
        "abstract": "The application retrieves untrusted data from persistent client-side storage and passes it to unsafe JavaScript execution sinks.",
        "impact": "Attackers may execute persistent malicious code in users’ browsers, leading to session theft, account compromise, or malicious browser actions."
    },

    "mvc view injection": {
        "abstract": "The application renders untrusted user input within MVC views without sufficient output encoding or validation.",
        "impact": "Attackers may inject malicious scripts or markup into rendered pages, potentially compromising user sessions or application integrity."
    },

    "parameter tampering": {
        "abstract": "The application trusts client-supplied parameters for security-sensitive operations without sufficient server-side validation.",
        "impact": "Attackers may manipulate request parameters to bypass authorization checks, alter transactions, or gain unauthorized access to resources."
    }
}

# ---------------------------------------------------------------------------
# Default Values
# ---------------------------------------------------------------------------

DEFAULT_ABSTRACT = (
    "The application contains a security weakness identified during "
    "the static source code review."
)

DEFAULT_IMPACT = (
    "Successful exploitation may impact the confidentiality, integrity, "
    "or availability of the application and associated data."
)

# ---------------------------------------------------------------------------
# Function to Match Vulnerability
# ---------------------------------------------------------------------------

def get_vuln_details(vuln_title):

    vuln_title = str(vuln_title).lower().strip()

    for keyword, details in VULN_DETAILS.items():

        if keyword in vuln_title:
            return details["abstract"], details["impact"]

    return DEFAULT_ABSTRACT, DEFAULT_IMPACT

# ---------------------------------------------------------------------------
# Read Excel File
# ---------------------------------------------------------------------------

print(f"[+] Reading file: {INPUT_FILE}")

# 4th sheet + headers on 3rd row
df = pd.read_excel(
    INPUT_FILE,
    sheet_name=3,
    header=2
)

# ---------------------------------------------------------------------------
# Clean Column Names
# ---------------------------------------------------------------------------

df.columns = (
    df.columns
    .astype(str)
    .str.replace("\n", " ", regex=False)
    .str.replace("\r", " ", regex=False)
    .str.strip()
)

print("\n[+] Detected Columns:")
print(df.columns.tolist())

# ---------------------------------------------------------------------------
# Detect Vulnerability Column
# ---------------------------------------------------------------------------

vuln_column = None

for col in df.columns:

    col_lower = col.lower().strip()

    if "vulnerability" in col_lower and "title" in col_lower:
        vuln_column = col
        break

if vuln_column is None:
    raise Exception("Could not find Vulnerability Title column")

print(f"\n[+] Using vulnerability column: {vuln_column}")

# ---------------------------------------------------------------------------
# Create / Replace Columns
# ---------------------------------------------------------------------------

df["Abstract"] = ""
df["Impact"] = ""

# ---------------------------------------------------------------------------
# Update Rows
# ---------------------------------------------------------------------------

for index, row in df.iterrows():

    vuln_title = row[vuln_column]

    abstract, impact = get_vuln_details(vuln_title)

    df.at[index, "Abstract"] = abstract
    df.at[index, "Impact"] = impact

# ---------------------------------------------------------------------------
# Create Output Folder
# ---------------------------------------------------------------------------

output_dir = os.path.dirname(OUTPUT_FILE)

if output_dir:
    os.makedirs(output_dir, exist_ok=True)

# ---------------------------------------------------------------------------
# Save Output
# ---------------------------------------------------------------------------

df.to_excel(OUTPUT_FILE, index=False)

print(f"\n[+] Updated file saved: {OUTPUT_FILE}")