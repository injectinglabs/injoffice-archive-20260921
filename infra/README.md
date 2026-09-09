# Public browser demo on AWS

`demo-site.yaml` manages the public `https://injoffice.com` demo. The GitHub
repository remains private. Public browser code includes the deliberately
bundled “View source” examples. No model, credentials, API bridge, or optional
Node/PDF/collaboration server is deployed.

## Infrastructure

Deploy the stack in **us-east-1** because CloudFront requires its ACM certificate
in that region. It creates a DNS-validated certificate, private encrypted and
versioned S3 bucket, CloudFront origin access control, cache/security policies,
and optional apex A/AAAA aliases in the existing hosted zone. HTTP redirects to
HTTPS. S3 public access is blocked; only the named distribution can read objects.

The deployment has two phases: create with `PublishDns=false`, upload and verify
the CloudFront endpoint, then update with `PublishDns=true`. Certificate DNS
validation happens during the first phase; the website aliases do not.

Application release: `50e0cbfbceaf944bbc679c828130505e8c7f418e` (PR #17).
Stack: `injoffice-demo-site`; AWS account: `707600960562`.
Hosted zone: `Z037397639ZG11Y7H1802`.
Bucket: `injoffice-demo-site-707600960562-us-east-1`.

## Build and publish

Use a clean checkout of a tested application commit, Node 22+ and the documented
Go/WASM toolchain. Keep environment files and credentials out of the build.

```sh
npm ci
npm run build -w apps/playground -- --base=/
aws cloudformation validate-template --region us-east-1 --template-body file://infra/demo-site.yaml
```

Review a CloudFormation change set before executing it. Pass `DomainName`, the
existing `HostedZoneId`, a full 40-character application commit as `ReleaseId`,
and `PublishDns`. Do not replace the existing hosted zone or its NS/SOA records.

Upload **only** the audited `apps/playground/dist` output:

1. Upload `dist/assets/` to bucket `assets/`, without deletion, with
   `Cache-Control: public,max-age=31536000,immutable`. These content-hashed paths
   are shared across releases, keeping old lazy chunks reachable by open tabs.
2. Set `.wasm` to `application/wasm`, `.js`/`.mjs` to `text/javascript`, and CSS to
   `text/css`; verify response types through CloudFront with `nosniff` enabled.
3. Upload non-asset files under `releases/<ReleaseId>/` with
   `Cache-Control: public,max-age=0,must-revalidate`. Upload `index.html` last.
4. Verify the release through the distribution's HTTPS hostname before enabling
   DNS. The app uses hash routes, so missing files must remain 403/404 responses,
   **not** an HTML SPA fallback.
5. For later releases/rollbacks, update `ReleaseId` only after uploading the full
   release, then invalidate `/*` and wait for completion. Origin paths do not
   change CloudFront viewer cache keys.
6. Run the existing-server browser checks using `SHOWCASE_URL=https://injoffice.com/`:
   `node scripts/smoke-showcase-browser.mjs` and
   `node scripts/smoke-scroll-showcase-browser.mjs`.

Do not use `s3 sync --delete`, upload repository source wholesale, overwrite a
previous release with different bytes, or delete older assets during a release.
Inspect the stack outputs for exact bucket/distribution identities.

## Rollback and maintenance

Keep older release prefixes and hashed assets. To roll back, restore the previous
tested `ReleaseId` through a reviewed stack update and invalidate `/*`. To take
the apex offline without deleting content, set `PublishDns=false`. The ACM
validation CNAME must remain for certificate renewal.

The bucket has retention policies: deleting or rolling back the stack does not
erase uploaded files. If initial creation rolls back, inspect retained resources
before retrying; do not delete the bucket to resolve a name conflict. Storage,
DNS and delivery incur normal AWS usage charges. Asset cleanup, logging/budgets,
and optional backend hosting require separate decisions.

This is an explicitly operated deployment. The existing private GitHub Pages
workflow does not update AWS; this setup grants no CI deployment credentials.

References: [AWS secure static hosting](https://docs.aws.amazon.com/AmazonCloudFront/latest/DeveloperGuide/getting-started-secure-static-website-cloudformation-template.html),
[private S3 origins](https://docs.aws.amazon.com/AmazonCloudFront/latest/DeveloperGuide/private-content-restricting-access-to-s3.html),
[CloudFront certificate requirements](https://docs.aws.amazon.com/AmazonCloudFront/latest/DeveloperGuide/cnames-and-https-requirements.html).
