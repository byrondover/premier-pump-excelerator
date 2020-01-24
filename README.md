# Premier Pump Excelerator

Open XML (Microsoft Excel 2007, etc.) spreadsheet manipulations for Premier Pump & Power, LLC.

http://www.premierpumpandpower.com

## Building

```shell
gcloud builds submit --tag gcr.io/premier-pump-excelerator/excelerator
```

## Deploying

```shell
gcloud builds submit --tag gcr.io/premier-pump-excelerator/excelerator
gcloud --project premier-pump-excelerator run deploy --image gcr.io/premier-pump-excelerator/excelerator --platform managed --max-instances 1 excelerator
```
