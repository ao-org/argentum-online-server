# Special fishing terminal safety

`modFishing.PerformFishing` owns selection and eligibility, marks the pending fish,
sends the existing payload-free server packet 176 and stops automatic fishing.
The existing client terminal packets are 286 (complete) and 287 (break rod/net).
No packet layout, probability, cost, reward selection or capability changes.

`HandleRomperCania` previously neither checked nor cleared `PescandoEspecial`.
Repeated failure messages could remove additional tools; a subsequent completion
could still award the pending catch. `TakeSpecialFishingCatch` now consumes the
pending catch before either terminal handler mutates inventory. No pending catch
means no reward/tool mutation. Dead users cannot complete. Walk, death and warp
clear the pending catch using `CancelSpecialFishing`; normal session teardown
continues to reset user state.

The existing protocol does not transmit a challenge token, score or timing proof.
This fix guarantees one terminal mutation per currently pending catch, not
anti-cheat verification or correlation across an entirely new challenge. HOO's
companion UI latches one result and suppresses duplicate starts until new fishing.
Deploy this correction before enabling the restored HOO minigame.

Suite 46 adds three cases through the real completion/break handlers: duplicate
completion, repeated break then completion, and cancelled/dead pending catches.
All databases/inventory are disposable harness state. Native VB6 is unavailable
on the local development machine; use PR validation for compilation and execution.
Manual checks must use disposable fishing tools because legacy failure intentionally
breaks a rod/removes a net. Verify success, early failure, timeout, cancellation,
death, movement, warp and restarting ordinary automatic fishing with HOO.
