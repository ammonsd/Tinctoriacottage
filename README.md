# Tinctoria Cottage

A vintage textiles website featuring handwoven items, vintage linens, and textile-related books.

## 🌐 Live Website

**[tinctoriacottage.com](http://tinctoriacottage.com)**

## 📖 About

Tinctoria Cottage is dedicated to sharing a love of textiles - from vintage linens and handcrafted items to carefully curated books about weaving, spinning, and textile arts. The site features:

- Vintage linens (table linens, hankies, doilies)
- Handwoven textiles and creations
- Vintage books on weaving, spinning, and textile crafts
- Curated links to weaving and textile resources

## 🏗️ Technical Details

### Hosting
- **Platform:** AWS S3 Static Website Hosting
- **Domain:** Route 53 DNS Management
- **Region:** US East (N. Virginia)

### Migration
This website was migrated from traditional web hosting to AWS S3 in November 2025, resulting in:
- **95% cost reduction** (from $20+/month to ~$0.54/month)
- Improved reliability and scalability
- Converted from SHTML/SSI to static HTML

### Technology Stack
- Static HTML/CSS/JavaScript
- Server-side includes converted to embedded HTML
- Legacy ASP files preserved in `/Dir` folder (not active)
- Image gallery and navigation system

## 📁 Repository Structure

```
├── index.html              # Homepage
├── linenpg.html           # Vintage linens page
├── creationpg.html        # Vintage creations page
├── wovenpg.html           # Handwovens page
├── bookpg.html            # Books page with order form
├── linkspg.html           # Textile resources links
├── graphics/              # Images and navigation elements
├── eBay/                  # Product photos
├── includes/              # Legacy SSI files (archived)
├── Dir/                   # Legacy ASP application (archived)
└── AWS-S3-Migration-Guide.md  # Complete migration documentation
```

## 🚀 Migration Guide

Want to migrate your own website to AWS S3? Check out the comprehensive guide:

**[AWS S3 Migration Guide](./AWS-S3-Migration-Guide.md)**

This detailed guide covers:
- Converting SHTML/SSI files to HTML
- Setting up S3 buckets and static hosting
- Configuring custom domains with Route 53
- Troubleshooting common issues
- Cost estimates and optimization

## 🔧 Local Development

To work with this repository locally:

```bash
# Clone the repository
git clone https://github.com/ammonsd/Tinctoriacottage.git

# Navigate to the directory
cd Tinctoriacottage

# Open in your preferred editor
code .
```

Since this is a static website, you can open `index.html` directly in a browser or use a local server:

```bash
# Python 3
python -m http.server 8000

# Or use VS Code Live Server extension
```

## 📤 Deployment

To deploy changes to S3:

```bash
# Using AWS CLI (after configuring credentials)
aws s3 sync . s3://tinctoriacottage.com --exclude ".git/*" --exclude "*.md"
```

Or upload files directly through the AWS S3 Console.

## 📝 History

- **November 2025:** Migrated to AWS S3 with custom domain
- **Legacy:** Originally hosted on traditional web server with ASP/SHTML
- **Conversion:** All SHTML files converted to HTML, SSI content embedded

## 📧 Contact

For inquiries about Tinctoria Cottage, please use the contact information on the website.

## 📄 License

All illustrations and photos are copyrighted and may not be used without prior permission.

---

*Website design preserves the original vintage aesthetic while modernizing the infrastructure for cost-effective, reliable hosting.*
