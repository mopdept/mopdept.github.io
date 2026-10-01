WASHROOM WATCH — Setup
1) Google Sheet খুলে Extensions > Apps Script > Code.gs পেস্ট করো। ADMIN_PIN বদলাও (চাইলে ADMIN_EMAIL দাও)।
2) Deploy > New deployment > Web app > Execute as: Me, Access: Anyone > /exec URL কপি করো।
3) Cloudinary > Settings > Upload > Unsigned preset বানাও।
4) index.html এ CONFIG (SCRIPT_URL, CLOUD_NAME, UPLOAD_PRESET, DEPARTMENTS, LOCATIONS) বসাও।
5) departments.html ও dashboard.html এ শুধু SCRIPT_URL বসাও।
6) সব ফাইল GitHub Pages / যেকোনো hosting এ আপলোড করো।
Pages: index.html (report + feed + admin), departments.html (dept-wise photos), dashboard.html (analytics + CSV).
