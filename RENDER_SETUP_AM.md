# Render ላይ የPastry Accounting Bot ማስኬጃ

1. ይህን folder ወደ GitHub repository ይጫኑ።
2. Render → New → Blueprint ይምረጡ።
3. Repository-ውን ይምረጡ፤ `render.yaml` በራሱ የWorker አገልግሎት ያዘጋጃል።
4. Secret Environment Variables:
   - `BOT_TOKEN` = የTelegram Bot Token
   - `ADMIN_IDS` = የAdmin Telegram numeric ID
5. Deploy ይጫኑ።

ማስታወሻ:
- ይህ bot Telegram polling ስለሚጠቀም Background Worker ነው።
- SQLite database እንዳይጠፋ `/var/data` persistent disk ተዘጋጅቷል።
- Render Background Workers በFree plan አይገኙም፤ ስለዚህ የRender plan መስፈርት አለው።
