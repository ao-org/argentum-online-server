# Priest Forgiveness Dialog

HOO uses the existing NPC interaction card for priest forgiveness. The server
continues to accept `/PERDON` and `/DONAR`; these legacy commands use the same
eligibility and checked price calculation as the dialog.

## Wire Contract

The server advertises the generic `hoo-protocol-v2` toggle. HOO acknowledges
protocol version 2 through the existing `eHooClientCapabilities` packet and
does not send the forgiveness request unless that toggle was advertised.
Version 1 capability masks and bits are unchanged. Older clients ignore the
unknown toggle, and the server continues accepting protocol version 1.

Client packet 320, `eHooPriestForgiveness`, is 13 bytes total.

| Field | Type | Meaning |
| --- | --- | --- |
| Packet ID | Int16 | 320 |
| Request ID | Int32 | Positive correlation ID |
| Action | UInt8 | 0 requests an offer; 1 confirms |
| Priest character index | Int16 | Must match the server-selected NPC |
| Offer token | Int32 | Zero for request; positive for confirmation |

Server packet 227, `eHooPriestForgivenessResult`, is 15 bytes total.

| Field | Type | Meaning |
| --- | --- | --- |
| Packet ID | Int16 | 227 |
| Request ID | Int32 | Echo of the request |
| Result | UInt8 | Result code below |
| Price | Int32 | Exact server-computed gold cost |
| Offer token | Int32 | Positive only for Offer/PriceChanged |

Result codes are: Offer=0, Success=1, AlreadyCitizen=2, Faction=3, Guild=4,
InsufficientGold=5, Dead=6, InvalidPriest=7, TooFar=8, InvalidOffer=9,
PriceChanged=10, InvalidPrice=11, RestrictedLocation=12, and
TransitionFailed=13.

An offer is bound to the authenticated user slot and session version, request
ID, exact NPC reference (including its lifetime version), and price. It expires
after 30 seconds using the server's wrap-safe monotonic tick helper. Confirm
rechecks the target, eligibility, current price, and carried gold. A changed
price returns a replacement offer; another explicit confirmation is required.
The offer is consumed before mutation, so duplicate confirmations cannot
charge twice. The server deducts the exact quoted price for HOO, updates gold,
and applies the existing citizenship transition and effects. If that transition
does not take effect, gold is restored and the client receives a rejection.

`VolverCiudadano` deliberately refuses citizenship transitions on trigger-6
tiles. The shared eligibility check now reports that restriction before a
donation is accepted. Faction values were checked against the server enums:
criminal is eligible; citizen/Armada/consejo are already citizens; Caos and
concilio cannot receive this forgiveness. Criminal and chaotic guilds remain
ineligible. The unreachable legacy `/PERDON` continuation was not revived.

## Rollout

Deploy the server before HOO. Existing clients continue using their original
commands and packets. HOO displays the action only after seeing
`hoo-protocol-v2`; against an older server, ordinary NPC interaction remains
available and the new action is not offered. Rollback can remove/disable the
HOO client first; the server's appended packets remain unused by older clients.
