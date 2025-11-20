# Recolección de plantas

Este servidor incluye soporte para objetos de tipo `otPlants` que se pueden recolectar con herramientas específicas.

## Configuración de las plantas
- **Tipo de objeto**: establezca `ObjType=10` en el archivo `.dat` del objeto, que corresponde a `e_OBJType.otPlants`.
- **Cantidad y regeneración**: utilice `VidaUtil` para indicar la cantidad total disponible y `TiempoRegenerar` para definir los segundos necesarios para restaurar la reserva tras agotarse.
- **Ítems entregados**:
  - Si la planta siempre entrega el mismo ítem, configure `HarvestItemIndex`, `HarvestMinAmount` y `HarvestMaxAmount`.
  - Para variantes aleatorias existen dos opciones equivalentes:
    - Definir `HarvestVariants` y cada `VariantN` con el formato `objIndex-cantidad-peso`. La cantidad es opcional (si queda vacía se usa el rango mínimo/máximo) y el peso controla la probabilidad relativa.
    - Indicar todas las variantes en el mismo `HarvestItemIndex`, separadas por `|`, `;` o `,`. Cada entrada utiliza el mismo formato `objIndex-cantidad-peso`, por ejemplo `HarvestItemIndex=596--90|595--10` reparte un 90 %/10 % entre los objetos 596 y 595 usando el rango `HarvestMin/Max`.

### Ejemplo de objeto planta

```
[OBJ5187]
Name=Flor de la vida
Texto=Flor de la vida
GrhIndex=12439
ObjType=10            ; e_OBJType.otPlants
Agarrable=1
TiempoRegenerar=60    ; segundos para recuperar la reserva
VidaUtil=10           ; cantidad total disponible antes de regenerar
HarvestItemIndex=605  ; índice del ítem que entrega (ajústalo según tus datos)
HarvestMinAmount=1
HarvestMaxAmount=3
```

Con esta configuración la planta entrega entre 1 y 3 unidades del objeto `605` cada vez que la recolectas. Para probabilidades múltiples basta con usar la forma abreviada `HarvestItemIndex=596--90|595--10`. Si omites los campos `Harvest*`, la planta quedará sin drops y no otorgará recompensas.

## Herramienta requerida
Las plantas solo pueden extraerse con herramientas de trabajo de subtipo `eHerbalismShears` (por ejemplo, unas tijeras de botánico). Asegúrese de crear un objeto `otWorkingTools` con `Subtipo=10` y equiparlo antes de intentar recolectar. Un ejemplo completo sería:

```
[OBJ5188]
Name=Tijera de botánico
Texto=Tijera de botánico
GrhIndex=44415
ObjType=18            ; e_OBJType.otWorkingTools
Subtipo=10            ; e_WorkingToolsSubType.eHerbalismShears
Valor=0
Crucial=0
Manejo=0              ; opcional, usar el valor adecuado para tu servidor
MinSkill=0            ; habilidad mínima necesaria para usar la herramienta
SkHerreria=0          ; mantenlo en 0 si no se consume durabilidad
SkCarpinteria=0
SkAlquimia=0
en_Name=Botanist Shears
en_Texto=Botanist Shears
pt_Name=Tesoura de botânico
pt_Texto=Tesoura de botânico
```

Puedes ajustar los campos de localización o requisitos según tus necesidades, pero es imprescindible que `ObjType` sea `18` y `Subtipo` se establezca en `10`.

### Cómo usar las tijeras en el juego
1. **Equipa las tijeras** (`OBJ5188` en el ejemplo) como herramienta principal desde tu inventario.
2. **Activa la habilidad de trabajo _Alquimia_** en la ventana de habilidades o con el acceso rápido habitual.
3. **Haz clic sobre la planta** (`otPlants`) que quieres recolectar estando a una casilla de distancia (no encima de ella).
4. Cada intento consume 5 puntos de energía (STA); al tener éxito, verás el mensaje de recolección y los ítems configurados en la planta se depositan directamente en tu inventario.

## Funcionamiento en el mapa
Al cargar el mapa, los objetos `otPlants` se inicializan como recursos fijos. Cada vez que un jugador recolecta, se descuenta la cantidad disponible y se programa la regeneración automática utilizando los valores definidos en el objeto.
