# EffectsOverTime Module Documentation

## 1. General Purpose

The `EffectsOverTime.bas` module is a comprehensive system designed to manage dynamic status effects, buffs, debuffs, and other game mechanics that persist for a duration or trigger under specific conditions. It handles the entire lifecycle of these "Effects Over Time" (EOTs) for both player characters (Users) and Non-Player Characters (NPCs).

Key functionalities include:

*   **Effect Creation:** Instantiating various types of effects based on predefined templates (likely loaded from data files).
*   **Effect Management:** Tracking all active effects in the game, associating them with their targets, and managing their lifetimes.
*   **Periodic Updates:** Regularly updating each active effect, allowing them to perform actions or check for expiration based on elapsed time.
*   **Event-Driven Callbacks:** Enabling effects to react to specific game events occurring to their target, such as using magic, attacking, taking damage, or changing terrain.
*   **Object Pooling:** Utilizing a pooling mechanism for effect objects to improve performance by reusing instances instead of frequently creating and destroying them.
*   **Stat Modification:** Applying and removing temporary modifications to character/NPC stats as part of an effect's behavior.

The system is designed to be extensible, with different effect behaviors encapsulated within specific classes that implement a common `IBaseEffectOverTime` interface.

## 2. Relationship with `Codigo/EffectOverTime/` Classes

The `EffectsOverTime.bas` module acts as a manager and factory for various effect classes, many of which are expected to reside in the `Codigo/EffectOverTime/` directory.

*   **Instantiation:** The private `InstantiateEOT` function within `EffectsOverTime.bas` is responsible for creating new objects of specific effect classes based on an `e_EffectOverTimeType` enum value. This includes classes like:
    *   `AttrackEffect.cls`
    *   `BonusDamageEffect.cls`
    *   `DelayedBlast.cls`
    *   `MultipleAttacks.cls`
    *   `ProtectEffect.cls`
    *   `TransformEffect.cls`
    *   `UnequipItem.cls`
    *   And others implied by `e_EffectOverTimeType` such as `UpdateHpOverTime`, `StatModifier`, `EffectProvoke`, `EffectProvoked`, `clsTrap`, `DrunkEffect`, `TranslationEffect`, `ApplyEffectOnHit`, `UpdateManaOverTime`, `ApplyEffectToParty`.
*   **Interface:** All these effect classes are expected to implement a common interface, referred to as `IBaseEffectOverTime`. This interface defines a contract for methods like `Setup`, `Update`, `OnRemove`, and properties to manage the effect's state, target, caster, duration, and callback registrations.
*   **Management:** The module stores instances of these effect objects in a global list (`ActiveEffects`) for updates and in per-character lists (`UserList().EffectOverTime`, `NpcList().EffectOverTime`) for targeted callbacks.

## 3. Main Public Subroutines and Functions

*   **`InitializePools()`**: Pre-populates pools of various effect objects to be reused, optimizing memory and performance.
*   **`UpdateEffectOverTime()`**: Iterates through all globally active EOTs, calls their `Update` method with the elapsed time, and handles their removal if they expire or are marked for removal.
*   **`CreateEffect(...)`**: Creates a new EOT instance based on a predefined `EffectIndex` (which references data likely from `EffectOverTime()` array), associates it with a source and target, and adds it to the update loop and the target's effect list.
*   **`CreateTrap(...)`**: Specialized factory function to create and manage instances of `clsTrap`.
*   **`CreateDelayedBlast(...)`**: Specialized factory function for `DelayedBlast` effects.
*   **`CreateUnequip(...)`**: Specialized factory function for `UnequipItem` effects.
*   **`AddEffectToUpdate(Effect As IBaseEffectOverTime)`**: Adds an existing effect object to the global update list.
*   **`AddEffect(EffectList As t_EffectOverTimeList, Effect As IBaseEffectOverTime)`**: Adds an effect object to a specified list (e.g., a character's list).
*   **`RemoveEffect(EffectList As t_EffectOverTimeList, Effect As IBaseEffectOverTime, ...)`**: Removes a specific effect from a list.
*   **`FindEffectOfTypeOnTarget(EffectList As t_EffectOverTimeList, TargetType As e_EffectType) As IBaseEffectOverTime`**: Searches a list for an effect of a given `e_EffectType`.
*   **`FindEffectOnTarget(CasterIndex, EffectList, EffectId) As IBaseEffectOverTime`**: Finds a specific effect instance on a target based on its `EffectId` and caster, respecting stacking rules.
*   **`ClearEffectList(EffectList As t_EffectOverTimeList, ...)`**: Removes multiple effects from a list, with optional filters.
*   **Callback Functions:**
    *   `TargetUseMagic`, `TartgetWillAtack`, `TartgetDidHit`, `TargetFailedAttack`, `TargetApplyDamageReduction`, `TargetWasDamaged`, `TargetWillAttackPosition`, `TargetUpdateTerrain`: These functions are called by other game systems when a relevant event occurs. They iterate through the effects on the target and invoke the corresponding method on effects that have registered for that event.
*   **`ChangeOwner(...)`**: Transfers an effect from one entity to another.
*   **`ConvertToClientBuff(buffType As e_EffectType) As e_EffectType`**: Maps internal effect types to simpler types for client display.
*   **`ApplyEotModifier(...)`**: Applies a set of stat modifications (from an `EffectStats` structure) to a target entity.
*   **`RemoveEotModifier(...)`**: Reverses the stat modifications applied by `ApplyEotModifier`.

## 4. Notable Dependencies

*   **Global Data Structures:**
    *   `UserList()`: Array of user data. Each user object contains a `t_EffectOverTimeList` named `EffectOverTime` to hold their current effects.
    *   `NpcList()`: Array of NPC data. Each NPC object also contains an `EffectOverTime` list.
    *   `EffectOverTime()`: A global array or collection, likely loaded from data files (e.g., `Effects.dat` or integrated into `Hechizos.dat`). This structure defines the static properties of each effect type (e.g., duration, stat changes, callback flags, stacking rules). Referenced by `EffectIndex` or `EffectTypeId`.
*   **Effect Classes (`Codigo/EffectOverTime/` and others):** Instances of classes implementing `IBaseEffectOverTime` are created and managed by this module.
*   **`IBaseEffectOverTime` Interface:** Defines the common methods and properties for all effect objects.
*   **Global Functions:** `GetTickCount`, `IsArrayInitialized`, `SetMask`, `TraceError`, `RunningInVB`.
*   **Character/NPC Stat Modules (Implicit):** Functions like `UpdateIncreaseModifier`, `WriteInmovilizaOK`, `SetStatusMask`, `UnsetStatusMask` are called to apply changes to character/NPC stats and states, implying interaction with modules that manage these aspects.
*   **Enumerations and Types:**
    *   `e_EffectOverTimeType`: Enum for different categories of EOTs (e.g., HealthModifier, ApplyModifiers, Provoke).
    *   `e_EffectType`: Enum for effect categories for client display or filtering (e.g., Buff, Debuff, Neutral).
    *   `e_ReferenceType`: Enum to distinguish between User and NPC targets/sources (`eUser`, `eNpc`).
    *   `e_EffectCallbackMask`: Enum defining flags for which game events an effect should respond to.
    *   `e_EOTTargetLimit`: Enum defining how effects stack or replace each other on a target.
    *   `t_EffectOverTimeList`: A structure likely containing an array of `IBaseEffectOverTime` objects and a count.
    *   `t_AnyReference`: A structure to hold a generic reference to a User or NPC.
    *   `t_EffectOverTime`: Structure holding the static properties of an effect type (loaded from data files).

This module provides a flexible and centralized way to handle a wide variety
of temporary and conditional effects within the game.Okay, I've created the Markdown documentation. Now I'll create the Mermaid component diagram in `Documentation/EffectsOverTime.mmd`.

This diagram will show `EffectsOverTime.bas` interacting with the effect classes in the `Codigo/EffectOverTime/` directory, and its dependencies on core game data like `UserList`, `NpcList`, and the `EffectOverTime()` data definitions.
