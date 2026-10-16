/* =========================================================================
   ZULÚ — Relojes · Datos del catálogo
   =========================================================================
   Este es el ÚNICO archivo que se edita para mantener el catálogo.
   - Para agregar un reloj: copia un bloque { ... } y cámbialo.
   - Para marcar como vendido: cambia  disponibilidad: "vendido"
   - Estados de venta válidos: "disponible" | "apartado" | "vendido"
   - Condición (estado): texto libre, ej. "Nuevo", "Como nuevo", "Excelente"
   - fotos: lista de URLs o rutas. Si está vacía [], se muestra un marcador.
   ========================================================================= */

const CONFIG = {
  marca: "Zulú",
  submarca: "Watch Co.",
  tagline: "La hora de todos los pilotos",
  intro:
    "Relojes seleccionados con criterio, desde la aviación. Cada pieza se revisa, " +
    "se documenta y se entrega con la confianza de un trato directo.",

  // ⚠️ CAMBIAR — ponme tu número real (formato internacional, sin + ni espacios).
  // México: 52 + 1 + 10 dígitos  →  ej. 5215512345678
  whatsapp: "5215500000000",

  // ⚠️ CAMBIAR / verificar que el @ esté libre
  instagram: "zulu.watches",

  // Moneda por defecto para mostrar precios
  moneda: "MXN",
  simbolo: "$",

  // Tu nombre como agente (aparece en contacto)
  agente: "Tu nombre",
};

const RELOJES = [
  {
    id: "navitimer-b01-43",
    marca: "Breitling",
    modelo: "Navitimer B01 Chronograph 43",
    referencia: "AB0138241B1A1",
    anio: 2021,
    caja_mm: 43,
    material: "Acero inoxidable",
    precio: 142000,
    estado: "Como nuevo",
    incluye: "Caja y papeles completos (full set)",
    disponibilidad: "disponible",
    destacado: true,
    nota_aviacion:
      "El reloj de piloto por excelencia: su regla de cálculo circular servía para " +
      "cálculos de navegación en vuelo.",
    descripcion:
      "Movimiento de manufactura B01 con cronógrafo y 70 h de reserva. Esfera azul, " +
      "bisel con regla de cálculo. Pieza icónica de la aviación, en condición impecable.",
    fotos: [],
  },
  {
    id: "speedmaster-moonwatch",
    marca: "Omega",
    modelo: "Speedmaster Professional Moonwatch",
    referencia: "310.30.42.50.01.001",
    anio: 2022,
    caja_mm: 42,
    material: "Acero inoxidable",
    precio: 185000,
    estado: "Excelente",
    incluye: "Caja y papeles completos (full set)",
    disponibilidad: "disponible",
    destacado: true,
    nota_aviacion:
      "El primer reloj usado en la Luna. Del cielo al espacio: aviación en su máxima expresión.",
    descripcion:
      "Calibre 3861 Master Chronometer, cristal de zafiro (Hesalite disponible bajo pedido). " +
      "El cronógrafo más legendario de la historia, listo para toda una vida de uso.",
    fotos: [],
  },
  {
    id: "gmt-master-ii-pepsi",
    marca: "Rolex",
    modelo: "GMT-Master II \"Pepsi\"",
    referencia: "126710BLRO",
    anio: 2023,
    caja_mm: 40,
    material: "Acero Oystersteel",
    precio: 395000,
    estado: "Nuevo",
    incluye: "Caja y papeles completos (full set)",
    disponibilidad: "apartado",
    destacado: false,
    nota_aviacion:
      "Diseñado originalmente para los pilotos de Pan Am: segunda zona horaria para cruzar husos.",
    descripcion:
      "Bisel Cerachrom azul y rojo, calibre 3285 con función GMT. El compañero de viaje " +
      "definitivo. Actualmente apartado — escríbeme para lista de espera o piezas similares.",
    fotos: [],
  },
  {
    id: "tudor-black-bay-gmt",
    marca: "Tudor",
    modelo: "Black Bay GMT",
    referencia: "M79830RB-0001",
    anio: 2022,
    caja_mm: 41,
    material: "Acero inoxidable",
    precio: 78000,
    estado: "Excelente",
    incluye: "Caja y papeles completos (full set)",
    disponibilidad: "disponible",
    destacado: false,
    nota_aviacion:
      "GMT de vocación aeronáutica con un valor difícil de igualar en su categoría.",
    descripcion:
      "Movimiento de manufactura MT5652 con GMT, bisel bicolor azul/rojo. La entrada ideal " +
      "al mundo de los relojes de viaje, con acabados de nivel superior.",
    fotos: [],
  },
  {
    id: "iwc-pilot-mark-xviii",
    marca: "IWC",
    modelo: "Pilot's Watch Mark XVIII",
    referencia: "IW327011",
    anio: 2020,
    caja_mm: 40,
    material: "Acero inoxidable",
    precio: 95000,
    estado: "Muy bueno",
    incluye: "Caja y papeles completos (full set)",
    disponibilidad: "vendido",
    destacado: false,
    nota_aviacion:
      "Diseño de instrumento de cabina puro: esfera negra, números legibles, cero distracción.",
    descripcion:
      "La definición del reloj de aviador moderno. Pieza vendida — tengo acceso a modelos " +
      "similares, escríbeme si buscas uno.",
    fotos: [],
  },
];

// Exponer a la página
window.ZULU = { CONFIG, RELOJES };
