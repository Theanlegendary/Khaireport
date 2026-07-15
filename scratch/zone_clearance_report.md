# Zone-wise Clearance Report (9:00 AM vs 2:00 PM)

This report tracks the **Bills Cleared** and **% Cleared** by Zone for each category, using the exact formulas from your example:
- `Bills Cleared = 9:00 AM - 2:00 PM` (Positive means bills decreased/cleared, negative means bills increased)
- `% Cleared = (Bills Cleared / 9:00 AM) * 100` (Positive means clearance rate, negative means accumulation rate)

### 📍 Pickup Table
| Zone | 9:00 AM | 2:00 PM | Bills Cleared | % Cleared |
| :--- | :---: | :---: | :---: | :---: |
| Zone 1 | 12 | 47 | **-35** | **-291.67%** |
| Zone 2 | 14 | 36 | **-22** | **-157.14%** |
| Zone 3 | 11 | 43 | **-32** | **-290.91%** |
| Zone 4 | 0 | 11 | **-11** | **0.00%** |
| Zone 5 | 12 | 17 | **-5** | **-41.67%** |
| **TOTAL ALL ZONE** | **49** | **154** | **-105** | **-214.29%** |

### 📍 Delivery Table
| Zone | 9:00 AM | 2:00 PM | Bills Cleared | % Cleared |
| :--- | :---: | :---: | :---: | :---: |
| Zone 1 | 68 | 140 | **-72** | **-105.88%** |
| Zone 2 | 34 | 63 | **-29** | **-85.29%** |
| Zone 3 | 24 | 32 | **-8** | **-33.33%** |
| Zone 4 | 16 | 15 | **1** | **6.25%** |
| Zone 5 | 7 | 42 | **-35** | **-500.00%** |
| **TOTAL ALL ZONE** | **149** | **292** | **-143** | **-95.97%** |

### 📍 Pending Table
| Zone | 9:00 AM | 2:00 PM | Bills Cleared | % Cleared |
| :--- | :---: | :---: | :---: | :---: |
| Zone 1 | 178 | 431 | **-253** | **-142.13%** |
| Zone 2 | 161 | 286 | **-125** | **-77.64%** |
| Zone 3 | 271 | 142 | **129** | **47.60%** |
| Zone 4 | 81 | 152 | **-71** | **-87.65%** |
| Zone 5 | 57 | 159 | **-102** | **-178.95%** |
| **TOTAL ALL ZONE** | **748** | **1170** | **-422** | **-56.42%** |

### 📍 Grand Total Table (All Categories Combined)
| Zone | 9:00 AM | 2:00 PM | Bills Cleared | % Cleared |
| :--- | :---: | :---: | :---: | :---: |
| Zone 1 | 258 | 618 | **-360** | **-139.53%** |
| Zone 2 | 209 | 385 | **-176** | **-84.21%** |
| Zone 3 | 306 | 217 | **89** | **29.08%** |
| Zone 4 | 97 | 178 | **-81** | **-83.51%** |
| Zone 5 | 76 | 218 | **-142** | **-186.84%** |
| **TOTAL ALL ZONE** | **946** | **1616** | **-670** | **-70.82%** |

