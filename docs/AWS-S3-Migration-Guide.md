# AWS S3 Static Website Migration Guide

Complete guide for migrating static websites from traditional hosting to AWS S3 with custom domain support.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Phase 1: Prepare Website Files](#phase-1-prepare-website-files)
3. [Phase 2: Create S3 Buckets](#phase-2-create-s3-buckets)
4. [Phase 3: Configure Custom Domain with Route 53](#phase-3-configure-custom-domain-with-route-53)
5. [Phase 4: Upload and Test](#phase-4-upload-and-test)
6. [Phase 5: Update Domain Registrar](#phase-5-update-domain-registrar)
7. [Troubleshooting](#troubleshooting)
8. [Cost Estimates](#cost-estimates)

---

## Prerequisites

### Required Accounts

-   AWS account with billing enabled
-   Domain name registered (any registrar: GoDaddy, Hostway, Namecheap, etc.)
-   Access to domain registrar account to update nameservers

### Required Knowledge

-   Basic understanding of DNS
-   Familiarity with AWS Console
-   Command line basics (for file operations)

### Tools Needed

-   Text editor for file modifications
-   PowerShell or Terminal
-   Web browser for AWS Console

---

## Phase 1: Prepare Website Files

### Step 1.1: Convert Server-Side Includes (SSI)

**Problem:** S3 static hosting doesn't support Server Side Includes (.shtml files with `<!--#include file="..."-->` directives).

**Solution:** Convert SSI files to plain HTML by embedding included content.

#### Identify Files to Convert

```powershell
# Find all .shtml files
Get-ChildItem -Recurse -Filter *.shtml

# Find all .ssi include files
Get-ChildItem -Recurse -Filter *.ssi
```

#### Manual Conversion Process

For each `.shtml` file:

1. **Locate SSI directives:**

    ```html
    <!--#include file="includes/header.ssi" -->
    ```

2. **Open the referenced `.ssi` file** and copy its entire content

3. **Replace the SSI directive** with the actual content:

    ```html
    <!-- Before -->
    <!--#include file="includes/header.ssi" -->

    <!-- After -->
    <div class="header">
        <h1>Welcome</h1>
        <!-- Full content from header.ssi -->
    </div>
    ```

4. **Save as `.html`** (keep original `.shtml` for reference)

#### Example Conversion

**Original file:** `linenpg.shtml`

```html
<body>
    <!--#include file="includes/linenpg.ssi" -->
</body>
```

**Converted file:** `linenpg.html`

```html
<body>
    <br />
    <!-- Content previously in linenpg.ssi is now embedded -->
</body>
```

### Step 1.2: Update Internal Links

After converting `.shtml` to `.html`, update all internal references.

#### PowerShell Batch Update

```powershell
# Update all .shtml references to .html in all HTML files
Get-ChildItem -Recurse -Include *.html,*.htm | ForEach-Object {
    (Get-Content $_.FullName -Raw) -replace '\.shtml', '.html' |
    Set-Content $_.FullName -NoNewline
}
```

#### Verify Changes

```powershell
# Search for any remaining .shtml references
Get-ChildItem -Recurse -Include *.html,*.htm | Select-String "\.shtml" |
Select-Object Path, LineNumber, Line
```

Should return **0 results** when complete.

### Step 1.3: Fix Case-Sensitive File Paths

**Problem:** S3 is case-sensitive (Linux-based), while Windows is not.

**Common Issues:**

-   `Graphics/image.gif` vs `graphics/image.gif`
-   `Images/photo.jpg` vs `images/photo.jpg`

#### Find Case Mismatches

```powershell
# Check for files with uppercase letters (may indicate issues)
Get-ChildItem -Recurse -File | Where-Object {
    $_.Name -cmatch '[A-Z]' -and
    $_.Extension -match '\.(gif|jpg|jpeg|png|css|js)$'
} | Select-Object FullName
```

#### Fix Image Path References

```powershell
# Example: Fix Graphics/ to graphics/
Get-ChildItem -Recurse -Include *.html,*.htm | ForEach-Object {
    (Get-Content $_.FullName -Raw) -replace 'Graphics/', 'graphics/' |
    Set-Content $_.FullName -NoNewline
}
```

#### Verify Directory Names Match References

Ensure your actual folder structure matches what's in your HTML:

-   If HTML says `src="graphics/logo.gif"`, folder must be named `graphics` (lowercase)
-   If folder is `Graphics`, update all HTML references to match

---

## Phase 2: Create S3 Buckets

### Step 2.1: Create Main Website Bucket

1. **Log into AWS Console** → Navigate to **S3**

2. **Click "Create bucket"**

3. **Bucket Configuration:**

    - **Bucket name:** `yourdomain.com` (must match domain exactly)
        - Example: `tinctoriacottage.com`
    - **Region:** Choose closest to your audience (e.g., `us-east-1`)
    - **Uncheck** "Block all public access"
    - **Acknowledge** the warning about public access
    - **Leave other settings as default**

4. **Click "Create bucket"**

### Step 2.2: Enable Static Website Hosting

1. **Click on your bucket name** to open it

2. **Go to "Properties" tab**

3. **Scroll to "Static website hosting"** section

4. **Click "Edit"**

5. **Configure:**

    - **Static website hosting:** Enable
    - **Hosting type:** Host a static website
    - **Index document:** `index.html`
    - **Error document:** `404.html` (optional, create if needed)

6. **Click "Save changes"**

7. **Note the S3 website endpoint URL** (e.g., `http://yourdomain.com.s3-website-us-east-1.amazonaws.com`)

### Step 2.3: Set Bucket Policy for Public Access

1. **Go to "Permissions" tab**

2. **Scroll to "Bucket policy"**

3. **Click "Edit"**

4. **Paste this policy** (replace `yourdomain.com` with your bucket name):

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Sid": "PublicReadGetObject",
            "Effect": "Allow",
            "Principal": "*",
            "Action": "s3:GetObject",
            "Resource": "arn:aws:s3:::yourdomain.com/*"
        }
    ]
}
```

5. **Click "Save changes"**

### Step 2.4: Create WWW Redirect Bucket

1. **Create second bucket** named `www.yourdomain.com`

    - Example: `www.tinctoriacottage.com`
    - Same region as main bucket
    - Can block public access (no policy needed)

2. **Enable Static Website Hosting:**

    - **Static website hosting:** Enable
    - **Hosting type:** Redirect requests for an object
    - **Target bucket:** `yourdomain.com` (your main bucket)
    - **Protocol:** `http`

3. **Click "Save changes"**

---

## Phase 3: Configure Custom Domain with Route 53

### Step 3.1: Create Hosted Zone

1. **Navigate to Route 53** in AWS Console

2. **Click "Hosted zones"** → **"Create hosted zone"**

3. **Configure:**

    - **Domain name:** `yourdomain.com`
    - **Type:** Public hosted zone

4. **Click "Create hosted zone"**

5. **Note the 4 nameservers** (you'll need these later):
    ```
    ns-1975.awsdns-54.co.uk
    ns-714.awsdns-25.net
    ns-1365.awsdns-42.org
    ns-449.awsdns-56.com
    ```

### Step 3.2: Create DNS Records for Root Domain

1. **Click "Create record"**

2. **Configure A Record for root domain:**

    - **Record name:** Leave blank (for root domain)
    - **Record type:** A
    - **Alias:** Toggle ON
    - **Route traffic to:**
        - Endpoint: "Alias to S3 website endpoint"
        - Region: Select your bucket's region
        - S3 endpoint: Select `yourdomain.com` from dropdown
    - **Routing policy:** Simple routing

3. **Click "Create records"**

### Step 3.3: Create DNS Records for WWW Subdomain

1. **Click "Create record"** again

2. **Configure A Record for www:**

    - **Record name:** `www`
    - **Record type:** A
    - **Alias:** Toggle ON
    - **Route traffic to:**
        - Endpoint: "Alias to S3 website endpoint"
        - Region: Select your bucket's region
        - S3 endpoint: Select `www.yourdomain.com` from dropdown
    - **Routing policy:** Simple routing

3. **Click "Create records"**

### Step 3.4: Verify DNS Records

Your hosted zone should now have:

-   **NS record** (4 nameservers)
-   **SOA record** (start of authority)
-   **A record** for root domain → S3 bucket
-   **A record** for www subdomain → S3 redirect bucket

---

## Phase 4: Upload and Test

### Step 4.1: Upload Files to S3

#### Option A: AWS Console (Small Sites)

1. **Open your main S3 bucket**

2. **Click "Upload"**

3. **Drag and drop** your entire website folder structure

4. **Click "Upload"**

#### Option B: AWS CLI (Recommended for Large Sites)

1. **Install AWS CLI:**

    ```powershell
    # Windows - using Chocolatey
    choco install awscli

    # Or download from: https://aws.amazon.com/cli/
    ```

2. **Configure AWS credentials:**

    ```bash
    aws configure
    # Enter: Access Key ID, Secret Access Key, Region, Output format
    ```

3. **Sync local files to S3:**
    ```bash
    aws s3 sync . s3://yourdomain.com --exclude ".git/*" --exclude "*.md"
    ```

### Step 4.2: Test with S3 Website Endpoint

1. **Open your S3 website endpoint URL** in browser:

    ```
    http://yourdomain.com.s3-website-us-east-1.amazonaws.com
    ```

2. **Verify:**

    - ✅ Homepage loads correctly
    - ✅ All images display
    - ✅ Navigation links work
    - ✅ All internal links work (no 404 errors)
    - ✅ CSS and JavaScript load properly

3. **Check browser console** (F12) for any errors

### Step 4.3: Test Common Issues

#### Missing Images?

-   Check file path case sensitivity
-   Verify image files uploaded
-   Check browser console for 404 errors

#### Broken Links?

-   Ensure all `.shtml` converted to `.html`
-   Check internal links point to correct files

#### CSS Not Loading?

-   Verify CSS file paths are correct
-   Check MIME types (should auto-detect)

---

## Phase 5: Update Domain Registrar

### Step 5.1: Update Nameservers at Registrar

**IMPORTANT:** This makes your custom domain point to AWS.

1. **Log into your domain registrar** (GoDaddy, Hostway, Namecheap, etc.)

2. **Find DNS/Nameserver settings** for your domain

3. **Select "Custom nameservers"** or "Use custom DNS"

4. **Replace existing nameservers** with your 4 Route 53 nameservers:

    ```
    ns-1975.awsdns-54.co.uk
    ns-714.awsdns-25.net
    ns-1365.awsdns-42.org
    ns-449.awsdns-56.com
    ```

5. **Save changes**

### Step 5.2: Wait for DNS Propagation

-   **Typical time:** 5 minutes to 48 hours
-   **Average:** 1-2 hours
-   **Fast cases:** Can be minutes

#### Check Propagation Status

**Option A: Online Tools**

-   Visit: https://www.whatsmydns.net
-   Enter your domain name
-   Select "A" record type
-   Should show your S3 endpoint IP

**Option B: Command Line**

```powershell
# Windows
nslookup yourdomain.com

# Should eventually show Route 53 nameservers
```

### Step 5.3: Test Custom Domain

1. **Visit your domain:** `http://yourdomain.com`

    - Should load your website from S3

2. **Test www subdomain:** `http://www.yourdomain.com`

    - Should redirect to root domain

3. **Test all pages and links**

---

## Troubleshooting

### Images Not Displaying

**Problem:** Images show broken icon

**Solutions:**

1. Check case sensitivity - `Graphics/` vs `graphics/`
2. Verify image files uploaded to correct S3 folders
3. Check browser console (F12) for 404 errors
4. Verify image paths in HTML match S3 structure

```powershell
# Find uppercase in image paths
Get-ChildItem -Recurse -Include *.html | Select-String -Pattern 'src="[A-Z]' -CaseSensitive
```

### 403 Forbidden Error

**Problem:** Can't access website, get "Access Denied"

**Solutions:**

1. **Check bucket policy** - ensure it allows public read
2. **Verify "Block public access"** is turned OFF
3. **Check file permissions** - ensure objects are readable
4. **Validate bucket policy** syntax (use JSON validator)

### 404 Not Found for Homepage

**Problem:** Domain loads but shows 404

**Solutions:**

1. **Verify `index.html` exists** in root of bucket
2. **Check static website hosting** is enabled
3. **Confirm index document** is set to `index.html`
4. **Test S3 endpoint URL** directly first

### DNS Not Resolving

**Problem:** Domain doesn't resolve after 24+ hours

**Solutions:**

1. **Verify nameservers** at registrar match Route 53 exactly
2. **Check Route 53 A records** point to correct S3 endpoints
3. **Clear DNS cache:**
    ```powershell
    # Windows
    ipconfig /flushdns
    ```
4. **Try different network** (mobile data, different WiFi)

### Mixed Content Warnings

**Problem:** Browser shows "not secure" warnings

**Solution:** Add CloudFront with SSL certificate (optional enhancement)

1. **Create CloudFront distribution** pointing to S3 bucket
2. **Request SSL certificate** via AWS Certificate Manager
3. **Update Route 53** A records to point to CloudFront (not S3)

### Website Slow to Load

**Solutions:**

1. **Enable CloudFront CDN** for faster global delivery
2. **Optimize images** - compress before uploading
3. **Choose S3 region** closest to majority of visitors

---

## Cost Estimates

### Typical Monthly Costs (Small Website)

| Service           | Description            | Typical Cost                   |
| ----------------- | ---------------------- | ------------------------------ |
| **S3 Storage**    | 1 GB of files          | $0.023/GB = **$0.02**          |
| **S3 Requests**   | 10,000 GET requests    | $0.0004 per 1,000 = **$0.004** |
| **Route 53**      | Hosted zone + queries  | $0.50 + minimal = **$0.51**    |
| **Data Transfer** | First 1 GB out is free | **$0.00**                      |
| **TOTAL**         | Per month              | **~$0.54**                     |

### Cost Comparison

| Hosting Type                | Monthly Cost     |
| --------------------------- | ---------------- |
| **Traditional Web Hosting** | $5 - $30/month   |
| **AWS S3 + Route 53**       | $0.50 - $2/month |
| **Savings**                 | **90-95%**       |

### Cost Optimization Tips

1. **Enable S3 Intelligent-Tiering** for infrequently accessed files
2. **Use CloudFront** only if needed (adds ~$0.085/GB transfer)
3. **Delete old versions** if versioning enabled
4. **Monitor usage** via AWS Cost Explorer

---

## Optional Enhancements

### Add HTTPS Support

**Requirement:** CloudFront distribution + SSL certificate

1. **Request SSL certificate** via AWS Certificate Manager (free)
2. **Create CloudFront distribution** with S3 origin
3. **Update Route 53** to point to CloudFront (not S3)
4. **Force HTTPS** in CloudFront settings

### Add 404 Error Page

1. **Create `404.html`** in website root
2. **Upload to S3 bucket**
3. **Set error document** in S3 static hosting settings to `404.html`

### Enable Access Logging

1. **Create separate S3 bucket** for logs
2. **Enable server access logging** on main bucket
3. **Analyze logs** to understand traffic patterns

### Set Up Website Analytics

**Option A: Google Analytics**

-   Add tracking code to all HTML pages

**Option B: AWS CloudWatch**

-   Enable S3 request metrics
-   View in CloudWatch dashboard

---

## Migration Checklist

### Pre-Migration

-   [ ] Backup current website files
-   [ ] Convert all SHTML/SSI files to HTML
-   [ ] Update internal links (.shtml → .html)
-   [ ] Fix case-sensitive file paths
-   [ ] Test locally (verify all links and images)

### AWS Setup

-   [ ] Create main S3 bucket (`yourdomain.com`)
-   [ ] Create www redirect bucket (`www.yourdomain.com`)
-   [ ] Enable static website hosting on both buckets
-   [ ] Set bucket policy for public read access
-   [ ] Upload all website files to main bucket
-   [ ] Test S3 website endpoint URL

### DNS Configuration

-   [ ] Create Route 53 hosted zone
-   [ ] Create A record for root domain
-   [ ] Create A record for www subdomain
-   [ ] Note all 4 nameservers

### Domain Registrar

-   [ ] Update nameservers at registrar
-   [ ] Wait for DNS propagation (15 mins - 48 hours)
-   [ ] Test custom domain URL

### Post-Migration

-   [ ] Test all pages load correctly
-   [ ] Verify all images display
-   [ ] Check all navigation links
-   [ ] Test forms (if any)
-   [ ] Check on multiple devices
-   [ ] Monitor AWS billing

### Cleanup (After Successful Migration)

-   [ ] Cancel old hosting service
-   [ ] Archive old SHTML files (optional)
-   [ ] Set up GitHub repository for version control
-   [ ] Document any custom configurations

---

## Additional Resources

### AWS Documentation

-   [S3 Static Website Hosting](https://docs.aws.amazon.com/AmazonS3/latest/userguide/WebsiteHosting.html)
-   [Route 53 Getting Started](https://docs.aws.amazon.com/Route53/latest/DeveloperGuide/getting-started.html)
-   [S3 Pricing Calculator](https://calculator.aws/)

### Tools

-   [AWS CLI Documentation](https://aws.amazon.com/cli/)
-   [DNS Propagation Checker](https://www.whatsmydns.net)
-   [Image Optimization Tools](https://tinypng.com)

### Useful Commands

```powershell
# Check Git status
git status

# Find files by extension
Get-ChildItem -Recurse -Filter *.shtml

# Search for text in files
Get-ChildItem -Recurse -Include *.html | Select-String "searchterm"

# Count files
(Get-ChildItem -Recurse -File).Count

# Get folder size
(Get-ChildItem -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
```

---

## Notes

-   This guide was created based on the successful migration of **tinctoriacottage.com**
-   Actual costs may vary based on traffic and file size
-   DNS propagation times vary by registrar and geographic location
-   Always test thoroughly before canceling old hosting
-   Keep backups of original files before making changes

**Migration Date:** November 2025  
**Guide Version:** 1.0

---

_For questions or issues, refer to AWS support documentation or community forums._
