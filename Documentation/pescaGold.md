# Configuración de `pescaGold.ini`

El servidor puede limitar qué peces aparecen en determinados mapas mediante el archivo `pescaGold.ini`, que debe ubicarse en la carpeta definida por `DatPath` (la misma donde se guardan `pesca.dat` y `RecursosEspeciales.dat`).

## Formato

```ini
[General]
Count=2

[Restriction1]
Maps=34
Fish=2152,3320
Replacement=3321

[Restriction2]
Maps=48,49
Fish=4000
Replacement=3321
```

- **Count**: cantidad de secciones `RestrictionN` presentes en el archivo.
- **Maps**: lista (separada por comas) de identificadores de mapa donde los peces listados están permitidos.
- **Fish**: lista de identificadores de objetos (peces) que sólo pueden pescarse en los mapas indicados.
- **Replacement**: identificador del objeto que reemplazará al pez restringido cuando se obtenga en un mapa distinto a los definidos en `Maps`. Si se omite o vale `0`, se utilizará el reemplazo configurado en `Configuracion.ini` (`ReplacementSpecialFish`).

Si `pescaGold.ini` no existe o no contiene restricciones válidas, el servidor utilizará los valores configurados en `Configuracion.ini` (campos `FishingMapSpecialFishID`, `UniqueMapfish1`, `UniqueMapfish2` y `ReplacementSpecialFish`).
