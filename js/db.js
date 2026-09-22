// ==========================================================
// AETHERIA AutoSend Pro — Persistent IndexedDB Database Engine
// Stores campaigns, contact lists, templates, and delivery logs
// ==========================================================

const DB_NAME = 'AetheriaAutoSendDB';
const DB_VERSION = 1;

class DatabaseEngine {
  constructor() {
    this.db = null;
    this.readyPromise = this.init();
  }

  init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (e) => {
        const db = e.target.result;

        // Store 1: Campaigns (file data, normalized rows, status)
        if (!db.objectStoreNames.contains('campaigns')) {
          const campStore = db.createObjectStore('campaigns', { keyPath: 'id' });
          campStore.createIndex('createdAt', 'createdAt', { unique: false });
          campStore.createIndex('name', 'name', { unique: false });
        }

        // Store 2: Saved Message Templates
        if (!db.objectStoreNames.contains('templates')) {
          const tplStore = db.createObjectStore('templates', { keyPath: 'id' });
          tplStore.createIndex('title', 'title', { unique: false });
        }

        // Store 3: Global Settings & App State
        if (!db.objectStoreNames.contains('settings')) {
          db.createObjectStore('settings', { keyPath: 'key' });
        }
      };

      request.onsuccess = (e) => {
        this.db = e.target.result;
        this.seedDefaultTemplates();
        resolve(this.db);
      };

      request.onerror = (e) => {
        console.error('IndexedDB open error:', e.target.error);
        reject(e.target.error);
      };
    });
  }

  async seedDefaultTemplates() {
    const existing = await this.getAllTemplates();
    if (existing.length === 0) {
      const defaults = [
        {
          id: 'tpl_invoice',
          title: 'Invoice Due Reminder',
          content: 'Hello {{Name}}, this is a friendly reminder that invoice #{{InvoiceNo}} of ₹{{Amount}} is due on {{DueDate}}. Please clear it at your earliest convenience. Thank you!',
          createdAt: new Date().toISOString()
        },
        {
          id: 'tpl_welcome',
          title: 'Welcome / Order Confirmation',
          content: 'Hello {{Name}}, thank you for your order with AETHERIA Solutions! Your order details: {{Details}}. We will notify you once dispatched.',
          createdAt: new Date().toISOString()
        },
        {
          id: 'tpl_followup',
          title: 'Customer Follow-up',
          content: 'Hi {{Name}}, following up on our recent conversation. Let us know if you need any assistance with your account. Have a great day!',
          createdAt: new Date().toISOString()
        }
      ];

      for (const t of defaults) {
        await this.saveTemplate(t);
      }
    }
  }

  // --- CAMPAIGN OPERATIONS ---

  async saveCampaign(campaign) {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['campaigns'], 'readwrite');
      const store = tx.objectStore('campaigns');

      if (!campaign.id) {
        campaign.id = 'camp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
      }
      campaign.updatedAt = new Date().toISOString();
      if (!campaign.createdAt) {
        campaign.createdAt = campaign.updatedAt;
      }

      const request = store.put(campaign);
      request.onsuccess = () => resolve(campaign);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async getCampaign(id) {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['campaigns'], 'readonly');
      const store = tx.objectStore('campaigns');
      const request = store.get(id);
      request.onsuccess = () => resolve(request.result);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async getAllCampaigns() {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['campaigns'], 'readonly');
      const store = tx.objectStore('campaigns');
      const request = store.getAll();
      request.onsuccess = () => {
        const sorted = (request.result || []).sort(
          (a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt)
        );
        resolve(sorted);
      };
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async deleteCampaign(id) {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['campaigns'], 'readwrite');
      const store = tx.objectStore('campaigns');
      const request = store.delete(id);
      request.onsuccess = () => resolve(true);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  // Update specific row status in active campaign
  async updateRowStatus(campaignId, rowIndex, status, logEntry = null) {
    await this.readyPromise;
    const campaign = await this.getCampaign(campaignId);
    if (!campaign) return null;

    if (campaign.data && campaign.data[rowIndex]) {
      campaign.data[rowIndex]._status = status;
    }

    if (logEntry) {
      if (!campaign.logs) campaign.logs = [];
      campaign.logs.push(logEntry);
    }

    // Recompute counts
    const sent = (campaign.data || []).filter(r => r._status === 'sent').length;
    const skipped = (campaign.data || []).filter(r => r._status === 'skipped').length;
    const failed = (campaign.data || []).filter(r => r._status === 'failed').length;
    campaign.stats = { sent, skipped, failed };

    return this.saveCampaign(campaign);
  }

  // --- TEMPLATE OPERATIONS ---

  async saveTemplate(template) {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['templates'], 'readwrite');
      const store = tx.objectStore('templates');

      if (!template.id) {
        template.id = 'tpl_' + Date.now();
      }
      if (!template.createdAt) {
        template.createdAt = new Date().toISOString();
      }

      const request = store.put(template);
      request.onsuccess = () => resolve(template);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async getAllTemplates() {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['templates'], 'readonly');
      const store = tx.objectStore('templates');
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async deleteTemplate(id) {
    await this.readyPromise;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(['templates'], 'readwrite');
      const store = tx.objectStore('templates');
      const request = store.delete(id);
      request.onsuccess = () => resolve(true);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  // --- BACKUP & RESTORE ---

  async exportFullBackup() {
    const campaigns = await this.getAllCampaigns();
    const templates = await this.getAllTemplates();
    return {
      exportedAt: new Date().toISOString(),
      version: 1,
      campaigns,
      templates
    };
  }

  async importFullBackup(backupData) {
    if (!backupData || !backupData.campaigns) {
      throw new Error('Invalid backup format');
    }

    for (const c of backupData.campaigns) {
      await this.saveCampaign(c);
    }
    if (backupData.templates) {
      for (const t of backupData.templates) {
        await this.saveTemplate(t);
      }
    }
    return true;
  }
}

// Global database instance
window.appDB = new DatabaseEngine();
