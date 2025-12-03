# lserm etfs
A utility to get ETFs (Exchange Traded Funds) out of [LSERM](https://www.ciro.ca/rules-and-enforcement/dealer-member-rules/supporting-schedules/list-securities-eligible-reduced-margin-lserm) (List of securities eligible for reduced margin) from <https://ciro.ca>. 

### Regex to detect missing commas:
- `"[^,a-zA-Z]`
- `[^=] "[a-zA-Z ]+"[^,a-zA-Z\]]`