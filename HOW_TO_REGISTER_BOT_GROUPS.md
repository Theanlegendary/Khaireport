# How to Register Groups for Bot Report Forwarding

## Overview
The bot can automatically forward daily reports to registered Telegram groups. This guide explains how to set up group registration.

---

## Method 1: Using /register Command (Recommended)

### Step-by-Step Instructions:

#### 1. **Add the Bot to Your Group**
- Open your Telegram group
- Click on the group name at the top
- Select "Add Members"
- Search for your bot's username
- Add the bot to the group
- Make sure the bot has permission to send messages

#### 2. **Register the Group**
Inside the group chat, send:
```
/register
```

#### 3. **Confirmation**
The bot will reply with:
```
✅ Group registered for report forwards.
Title: [Your Group Name]
Chat ID: [Group Chat ID]

Use /pause to pause forwarding and /resume to re-enable.
```

#### 4. **Auto-Detection (Smart Feature)**
The bot automatically detects post office handles from your group title!
- Example: If your group is named "PNPP014 Team", the bot will send PNPP014 reports
- Example: If your group is named "KAN Zone", the bot will send KANP001 reports

---

## Method 2: Manual Configuration (Advanced)

### Edit config.json

If you prefer manual configuration or need more control:

#### 1. **Open config.json**
```json
{
  "telegram": {
    "forward_mapping": {
      "CHAT_ID_HERE": ["PNPP001", "PNPP002"],
      "ANOTHER_CHAT_ID": ["*"]
    }
  }
}
```

#### 2. **Add Your Group**
- Replace `CHAT_ID_HERE` with your group's Chat ID
- Add the post office handles you want that group to receive
- Use `["*"]` to receive all reports

#### 3. **Finding Chat ID**
To find your group's Chat ID:
1. Add the bot to the group
2. Send `/register` in the group
3. The bot will show the Chat ID
4. Use `/unregister` if you don't want to keep it registered

---

## Managing Registered Groups

### View All Registered Groups
Send this command anywhere (bot will reply privately):
```
/groups
```

**Response:**
```
Registered groups:
  • PNPP001 Team — -1001234567890
  • KAN Zone — -1001234567891
  • All Reports — -1001234567892
```

### Unregister a Group
Inside the group you want to remove:
```
/unregister
```

---

## Forwarding Behavior

### What Gets Forwarded:

When you run the `push` command, registered groups receive:

#### For Specific Handle Groups:
- 📸 Report images for their assigned handles
- 📄 Excel files (if more than 50 pending orders)
- 💬 Remark messages with summary

#### For "*" (All Reports) Groups:
- 📸 All report images
- 📊 Zone summary reports
- 📄 Complete Excel file with all data
- 💬 Overall summary caption

### Smart Handle Detection:

The bot looks for handle codes in group titles:
- **Full Handle Format**: `PNPP014`, `KANP001`, `SVAP001`
- **Branch Code Only**: `SIH`, `KOH`, `PNP` → automatically converts to `[CODE]P001`

Examples:
- Group: "PNPP014 Daily Reports" → Receives PNPP014 data
- Group: "SIH Team" → Receives SIHP001 data
- Group: "Zone 1 Updates" → No auto-detection (use manual config)

---

## Control Commands

### Pause/Resume Forwarding

#### Pause All Forwarding:
```
/pause
```
- Bot still works for you personally
- No reports sent to any groups
- Good for maintenance or testing

#### Resume Forwarding:
```
/resume
```
- Resumes sending reports to all registered groups

#### Check Status:
```
/status
```
Shows whether forwarding is paused or active.

---

## Zone-Based Forwarding

### Registering Zone Groups

For zone-specific reports, add to `config.json`:

```json
{
  "zone_forward_mapping": {
    "ZONE_GROUP_CHAT_ID": "zone1"
  },
  "total_zones": {
    "zone1": ["PNPP001", "PNPP002", "PNPP003", "..."],
    "zone2": ["SPEP001", "TAKP001", "KAMP001", "..."],
    "zone3": ["BANP001", "BATP001", "PURP001", "..."],
    "zone4": ["CHHP001", "SIEP001", "PRHP001", "..."],
    "zone5": ["TBKP001", "CHAP001", "KRAP001", "..."]
  }
}
```

Zone groups receive zone-specific summary reports.

---

## Troubleshooting

### Group Not Receiving Reports?

**Check:**
1. ✅ Is the bot still in the group?
2. ✅ Does the bot have send message permissions?
3. ✅ Is forwarding paused? (Check with `/status`)
4. ✅ Is the group registered? (Check with `/groups`)
5. ✅ Run `/register` again to refresh

### Bot Can't Send to Group?

**Possible Issues:**
- Bot was removed from group → Add it back
- Group migrated to supergroup → Run `/register` again (bot auto-updates)
- Bot lacks permissions → Give "Send Messages" permission

### Want to Test First?

Use the `/test` command instead of `push`:
```
/test
```
- Fetches data and generates reports
- Sends ONLY to you (not to groups)
- Perfect for testing before going live

---

## File Storage

### Registered Groups File:
```
registered_groups.json
```
Location: Project root directory

Format:
```json
[
  {
    "chat_id": -1001234567890,
    "title": "PNPP001 Team"
  },
  {
    "chat_id": -1001234567891,
    "title": "All Reports"
  }
]
```

---

## Summary

### Quick Start:
1. Add bot to your Telegram group
2. Send `/register` in the group
3. Run `push` to send reports
4. Use `/pause` and `/resume` to control forwarding

### Best Practices:
- ✅ Include handle codes in group titles for auto-detection
- ✅ Use separate groups for different zones/handles
- ✅ Create one "All Reports" group with `"*"` mapping for managers
- ✅ Test with `/test` before using `push`
- ✅ Use `/groups` to verify registration

### Need Help?
- `/help` - View all available commands
- `/status` - Check bot and forwarding status
- `/groups` - List all registered groups

---

## Example Workflow

```bash
# In your group chat:
User: /register
Bot: ✅ Group registered for report forwards.
     Title: PNPP014 Daily Reports
     Chat ID: -1001234567890

# Later, when you run reports:
You (in PM): push
Bot: [Generates reports]
     → Sends to you
     → Forwards PNPP014 reports to registered group

# Check what groups are registered:
You: /groups
Bot: Registered groups:
     • PNPP014 Daily Reports — -1001234567890
     • All Reports — -1001234567899
```

---

Ready to register your groups! 🚀
