# Bot Commands Cheatsheet

## 📋 Group Registration Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/register` | Inside a group | Register the group to receive reports |
| `/unregister` | Inside a group | Remove the group from receiving reports |
| `/groups` | Anywhere | List all registered groups (bot replies privately) |

## 🚀 Report Generation Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `push` | Anywhere | Generate reports and send to YOU + all registered groups |
| `/test` | Anywhere | Generate reports but send ONLY to you (not groups) |
| `/dailyreport` | Anywhere | Generate Master Daily Report with all analytics |
| `/dailyreport 11/09` | Anywhere | Generate report for specific date (DD/MM) |

## ⏸️ Control Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/pause` | Anywhere | Pause forwarding to groups (bot still works for you) |
| `/resume` | Anywhere | Resume forwarding to groups |
| `/status` | Anywhere | Show if forwarding is paused/active |

## 🔍 Query Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/find [ORDER_ID]` | Anywhere | Search for an order by ID |
| `/check [ORDER_ID]` | Anywhere | Check order status |
| `/trace [ORDER_ID]` | Anywhere | Trace order history |
| `/total` | Anywhere | Show total orders summary |

## 📊 Comparison Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/vs [handle]` | Anywhere | Compare 8AM vs 2PM reports |
| `/vs2 [handle]` | Anywhere | Compare 2PM vs 5PM reports |

## ⚙️ Settings Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/mode` | Anywhere | Toggle between wide/long image mode |
| `/help` | Anywhere | Show available commands |

## 🗑️ Cleanup Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/clean` | Anywhere | Delete today's temporary report files |
| `/deletereport` | Anywhere | Delete specific report files |

## 📝 Delayed Bills Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/delay [ORDER_ID]` | Anywhere | Mark order as delayed |
| `/undelay [ORDER_ID]` | Anywhere | Remove delayed mark from order |
| `/delaylist` | Anywhere | Show all delayed orders |

## 🔢 Management Commands

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/add [ORDER_ID]` | Anywhere | Add order to tracking |
| `/remove [ORDER_ID]` | Anywhere | Remove order from tracking |
| `/list` | Anywhere | List tracked orders |

## 📱 QR Code

| Command | Where to Use | What It Does |
|---------|-------------|--------------|
| `/qr` | Anywhere | Generate QR code for web dashboard |

---

## 💡 Usage Examples

### Registering Groups
```
# In your group chat named "PNPP014 Team":
/register

# Bot responds:
✅ Group registered for report forwards.
Title: PNPP014 Team
Chat ID: -1001234567890
```

### Testing Before Going Live
```
# Test without sending to groups:
/test

# When ready to broadcast:
push
```

### Checking Registration Status
```
# See all registered groups:
/groups

# Response:
Registered groups:
  • PNPP014 Team — -1001234567890
  • All Reports — -1001234567891
```

### Generating Daily Reports
```
# Today's report:
/dailyreport

# Specific date:
/dailyreport 11/09
/dailyreport 11/09/2026
```

### Comparing Reports
```
# Morning vs Afternoon for specific handle:
/vs PNPP014

# Afternoon vs Evening:
/vs2 PNPP014

# For your zone (if you're in a zone group):
/vs
/vs2
```

### Pausing Forwarding
```
# Before testing or maintenance:
/pause

# Bot responds:
⏸ Forwarding to groups is now PAUSED.
The bot will still respond to your commands.

# Resume when ready:
/resume

# Bot responds:
▶️ Forwarding to groups is now ACTIVE.
```

### Finding Orders
```
# Search by order ID:
/find KH2609110001

# Check status:
/check KH2609110001

# Trace history:
/trace KH2609110001
```

---

## 🎯 Common Workflows

### Daily Morning Workflow
```
1. push                    # Generate and send morning reports
2. /vs                     # Later: Compare morning vs afternoon
3. /dailyreport           # End of day: Generate full report
```

### Setup New Group
```
1. Add bot to group
2. /register              # In the group
3. /groups                # Verify registration
4. /test                  # Test without broadcasting
5. push                   # Go live
```

### Maintenance Mode
```
1. /pause                 # Stop forwarding to groups
2. /test                  # Test changes
3. /resume                # Resume normal operation
```

---

## 📌 Pro Tips

✅ **Use descriptive group names** with handle codes (e.g., "PNPP014 Daily Reports")
✅ **Test first** with `/test` before using `push`
✅ **Use /pause** during maintenance or testing
✅ **Check /groups** regularly to verify registrations
✅ **Use /status** to see if forwarding is active

🚫 **Don't forget** to give the bot "Send Messages" permission in groups
🚫 **Don't remove** the bot from groups without unregistering first

---

## Quick Reference Card

```
╔════════════════════════════════════════╗
║  🎯 MOST USED COMMANDS                 ║
╠════════════════════════════════════════╣
║  /register    → Register group         ║
║  push         → Send reports           ║
║  /test        → Test only              ║
║  /dailyreport → Full daily report      ║
║  /groups      → List groups            ║
║  /pause       → Pause forwarding       ║
║  /resume      → Resume forwarding      ║
║  /status      → Check status           ║
╚════════════════════════════════════════╝
```
