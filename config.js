/* Подесувања за „Откуп на терен“ — ова е ЕДИНСТВЕНАТА датотека што се менува при инсталација.
   Вредностите се добиваат од Entra ID (регистрација на апликација) и од SharePoint. Види README.md. */
window.OTKUP_CONFIG = {
  // Entra ID (Azure AD) — регистрација на апликацијата
  tenantId: "4de6f642-d0ce-4c2e-983b-6c8ee45338aa",                 // Directory (tenant) ID
  clientId: "d4e71535-445a-42dd-afc5-5f68097a8d49",                 // Application (client) ID
  // SharePoint сајт каде стојат листите (host:/sites/име)
  site: "bfgj.sharepoint.com:/sites/mlekarnica",
  // Имиња на листите во SharePoint (ги создава setup.html)
  listaKooperanti: "Kooperanti",
  listaVnesovi: "OtkupVnesovi",
  // Текст во заглавието
  firma: "Ѓоргиеви ДООЕЛ",
  // Верзија — зголеми ја при секоја промена на датотеките, за таблетите да ја повлечат новата
  verzija: "2026-10-08a"
};
