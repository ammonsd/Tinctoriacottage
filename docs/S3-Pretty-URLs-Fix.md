# S3 Pretty URLs - Routing Rules Fix

## Problem Description

When accessing pages on the S3-hosted website, users encountered **404 errors** when:

-   Refreshing the page (F5)
-   Manually entering URLs like `http://tinctoriacottage.com/vintage-creations`
-   Accessing the actual file URLs like `http://tinctoriacottage.com/searchpg.html`

### Error Messages

```
404 Not Found
Code: NoSuchKey
Message: The specified key does not exist.
Key: books (or other path)
```

In some cases, an **infinite redirect loop** occurred:

```
ERR_TOO_MANY_REDIRECTS
This page redirected you too many times.
```

## Root Cause

Each HTML page contains JavaScript code that rewrites the browser's address bar to display "pretty URLs" without the `.html` extension:

```javascript
// Example from creationpg.html
document.addEventListener("DOMContentLoaded", function () {
    if (globalThis.location.pathname !== "/vintage-creations") {
        globalThis.history.replaceState(null, null, "/vintage-creations");
    }
});
```

### The Problem Flow:

1. User clicks link to `/creationpg.html`
2. Page loads successfully
3. JavaScript changes browser URL to `/vintage-creations` (pretty URL)
4. User refreshes the page
5. Browser requests `/vintage-creations` from S3
6. S3 treats this as a folder path (looks for `/vintage-creations/index.html`)
7. File doesn't exist → **404 error**

## Solution: S3 Routing Rules

Configure S3 to redirect pretty URLs back to their corresponding HTML files **only when they result in 404 errors**.

### Key Components:

1. **Routing Rules** - Map pretty URLs to actual files
2. **404 Error Condition** - Only redirect on missing files (prevents redirect loops)
3. **Error Document** - Fallback page for truly invalid URLs

## Implementation Steps

### 1. Create S3 Website Configuration

Create `s3-website-configuration.json` with routing rules that match each page's JavaScript pretty URL:

```json
{
    "IndexDocument": {
        "Suffix": "index.html"
    },
    "ErrorDocument": {
        "Key": "error.html"
    },
    "RoutingRules": [
        {
            "Condition": {
                "KeyPrefixEquals": "books",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "bookpg.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "home",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "index.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "vintage-creations",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "creationpg.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "vintage-linens",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "linenpg.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "handwovens",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "wovenpg.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "search",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "searchpg.html",
                "HttpRedirectCode": "301"
            }
        },
        {
            "Condition": {
                "KeyPrefixEquals": "favorite-links",
                "HttpErrorCodeReturnedEquals": "404"
            },
            "Redirect": {
                "ReplaceKeyWith": "linkspg.html",
                "HttpRedirectCode": "301"
            }
        }
    ]
}
```

### 2. Apply Configuration to S3 Bucket

#### Via AWS Console (Recommended):

1. Go to **S3 Console** → Select your bucket (e.g., `tinctoriacottage.com`)
2. Click **Properties** tab
3. Scroll to **Static website hosting** section
4. Click **Edit**
5. Configure:
    - **Static website hosting:** Enable
    - **Index document:** `index.html`
    - **Error document:** `error.html`
6. In **Redirection rules** section, paste the `RoutingRules` array (just the array, not the entire JSON)
7. Click **Save changes**

#### Via AWS CLI (Requires proper IAM permissions):

```bash
aws s3api put-bucket-website --bucket tinctoriacottage.com --website-configuration file://s3-website-configuration.json
```

### 3. Create Error Page

Create `error.html` for handling truly invalid URLs:

```html
<html>
    <head>
        <title>Page Not Found - Tinctoria Cottage</title>
        <meta
            http-equiv="Content-Type"
            content="text/html; charset=iso-8859-1"
        />
        <meta http-equiv="refresh" content="3;url=/index.html" />
    </head>
    <body bgcolor="#FFFFFF">
        <table
            width="100%"
            height="100%"
            border="0"
            cellpadding="0"
            cellspacing="0"
        >
            <tr>
                <td align="center" valign="middle">
                    <p>
                        <font face="Times New Roman, Times, serif" size="5"
                            >Page Not Found</font
                        >
                    </p>
                    <p>
                        <font face="Times New Roman, Times, serif"
                            >The page you requested could not be found.</font
                        >
                    </p>
                    <p>
                        <font face="Times New Roman, Times, serif"
                            >You will be redirected to the
                            <a href="/index.html">home page</a> in 3
                            seconds...</font
                        >
                    </p>
                </td>
            </tr>
        </table>
    </body>
</html>
```

Upload `error.html` to the root of your S3 bucket.

### 4. Update Modified Files

Upload the updated `searchpg.html` (and any other files modified during troubleshooting) to S3.

## URL Mapping Reference

| Pretty URL           | Actual File       | Page Description  |
| -------------------- | ----------------- | ----------------- |
| `/books`             | `bookpg.html`     | Books page        |
| `/home`              | `index.html`      | Home page         |
| `/vintage-creations` | `creationpg.html` | Vintage Creations |
| `/vintage-linens`    | `linenpg.html`    | Vintage Linens    |
| `/handwovens`        | `wovenpg.html`    | Handwovens page   |
| `/search`            | `searchpg.html`   | Search page       |
| `/favorite-links`    | `linkspg.html`    | Links page        |

## Testing

After applying the configuration:

1. **Test actual file URLs:**

    - `http://tinctoriacottage.com/bookpg.html` → Should load normally, URL changes to `/books`

2. **Test pretty URLs:**

    - `http://tinctoriacottage.com/books` → Should redirect to `bookpg.html`, then URL changes to `/books`

3. **Test refresh:**

    - Navigate to any page, press F5 → Should reload successfully

4. **Test invalid URLs:**
    - `http://tinctoriacottage.com/invalid-page` → Should show `error.html` and redirect to home

## Important Notes

### Why the 404 Condition is Critical

**Without** `"HttpErrorCodeReturnedEquals": "404"`:

-   `/search` → redirects to `searchpg.html` ✓
-   `/searchpg.html` → ALSO matches "search" prefix → redirects to itself → **infinite loop** ✗

**With** `"HttpErrorCodeReturnedEquals": "404"`:

-   `/search` (doesn't exist) → 404 → redirects to `searchpg.html` ✓
-   `/searchpg.html` (exists) → loads normally, no redirect ✓

### Adding New Pages

When adding a new page with a pretty URL:

1. Add the `history.replaceState` JavaScript to the HTML file
2. Add a corresponding routing rule to `s3-website-configuration.json`
3. Update the S3 bucket configuration
4. Upload the new HTML file to S3

### CloudFront Considerations

If you add CloudFront for HTTPS support in the future, you'll need to configure CloudFront's error pages to handle 404s similarly, or CloudFront may cache 404 responses before S3 routing rules can redirect.

## Common Issues & Solutions

### Issue: "Too Many Redirects"

**Cause:** Routing rule matches both the pretty URL and the actual filename  
**Solution:** Add `"HttpErrorCodeReturnedEquals": "404"` to the condition

### Issue: 404 Still Occurs After Update

**Cause:** DNS/CDN caching  
**Solution:**

-   Clear browser cache
-   Wait 5-10 minutes for S3 configuration to propagate
-   Test in incognito/private browsing mode

### Issue: Error Page Not Displaying

**Cause:** `error.html` not uploaded to S3 or error document not configured  
**Solution:**

-   Upload `error.html` to bucket root
-   Verify "Error document" is set to `error.html` in S3 static website hosting settings

## Files Modified/Created

-   ✅ `s3-website-configuration.json` - Created (S3 routing configuration)
-   ✅ `error.html` - Created (custom error page)
-   ✅ `searchpg.html` - Modified (changed `/product-search` to `/search`)
-   ✅ `docs/S3-Pretty-URLs-Fix.md` - Created (this document)
-   ❌ `s3-routing-rules.json` - Can be deleted (redundant, not used)

## Related Documentation

-   [AWS S3 Migration Guide](../AWS-S3-Migration-Guide.md)
-   [AWS S3 Static Website Hosting](https://docs.aws.amazon.com/AmazonS3/latest/userguide/WebsiteHosting.html)
-   [AWS S3 Routing Rules](https://docs.aws.amazon.com/AmazonS3/latest/userguide/how-to-page-redirect.html#advanced-conditional-redirects)

---

_Last Updated: December 17, 2025_
