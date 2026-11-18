import "@fontsource/inter/300.css";
import "@fontsource/inter/400.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/800.css";
import "@fontsource/jetbrains-mono/400.css";
import "@fontsource/jetbrains-mono/500.css";
import { getLength, getPointAtLength } from "@remotion/paths";
import { geoMercator, geoPath } from "d3-geo";
import { curveCatmullRom, line } from "d3-shape";
import type { Feature, FeatureCollection, Geometry } from "geojson";
import { feature } from "topojson-client";
import type { GeometryCollection, Topology } from "topojson-specification";
import countries50m from "world-atlas/countries-50m.json";
import {
  AbsoluteFill,
  Audio,
  Easing,
  interpolate,
  random,
  spring,
  staticFile,
  useCurrentFrame,
  useVideoConfig,
} from "remotion";

// Paleta nocturna, estilo cinematográfico
const BG = "#050A15";
const LAND = "#0F1C33";
const LAND_DIM = "#0A1426";
const BORDER = "#22406B";
const ACCENT = "#4EA8FF";
const GOLD = "#FFB547";
const WHITE = "#F4F7FB";
const MUTED = "#8DA2C0";
const SANS = "'Inter', sans-serif";
const MONO = "'JetBrains Mono', monospace";

const W = 1920;
const H = 1080;

// Momentos clave (en fotogramas, 30 fps)
const T = {
  titleOut: 80,
  routeStart: 105,
  routeEnd: 200,
  zoomStart: 215,
  zoomEnd: 260,
  step1: 258,
  step2: 303,
  step3: 348,
  closingIn: 395,
};

const ease = Easing.inOut(Easing.cubic);
const clamp = {
  extrapolateLeft: "clamp",
  extrapolateRight: "clamp",
} as const;

type LonLat = [number, number];
const MONTERREY: LonLat = [-100.31, 25.67];
const GUATE: LonLat = [-90.51, 14.63];
const XELA: LonLat = [-91.52, 14.84];
const ESCUINTLA: LonLat = [-90.79, 14.3];
// Ruta aproximada por carretera: Monterrey → CDMX → Oaxaca → Tapachula → Guatemala
const ROUTE: LonLat[] = [
  MONTERREY,
  [-100.98, 22.15],
  [-99.13, 19.43],
  [-97.6, 18.4],
  [-96.72, 17.06],
  [-94.5, 16.3],
  [-92.26, 14.9],
  GUATE,
];

const SHOWN = ["484", "320", "084", "340", "222", "840", "558", "192"];

const topo = countries50m as unknown as Topology<{
  countries: GeometryCollection<{ name: string }>;
}>;
const allCountries = feature(topo, topo.objects.countries) as FeatureCollection<
  Geometry,
  { name: string }
>;
const shapes = allCountries.features.filter((f) =>
  SHOWN.includes(String(f.id)),
) as Feature<Geometry, { name: string }>[];

const projection = geoMercator().fitExtent(
  [
    [0, 0],
    [W, H],
  ],
  {
    type: "Feature",
    properties: {},
    geometry: {
      type: "MultiPoint",
      coordinates: [
        [-118, 32],
        [-83, 12],
      ],
    },
  },
);
const pathGen = geoPath(projection);
const project = (p: LonLat) => projection(p) as [number, number];

const routeD = line().curve(curveCatmullRom.alpha(0.5))(ROUTE.map(project))!;
const routeLength = getLength(routeD);

const ZOOM = 3.8;
const focus = project([-90.5, 15.2]);

// Cámara: zoom en 2D dentro del plano + inclinación 3D del plano
const useCamera = () => {
  const frame = useCurrentFrame();
  const p = interpolate(frame, [T.zoomStart, T.zoomEnd], [0, 1], {
    ...clamp,
    easing: ease,
  });
  const k = Math.exp(interpolate(p, [0, 1], [0, Math.log(ZOOM)]));
  const sx = interpolate(p, [0, 1], [focus[0], W * 0.66]);
  const sy = interpolate(p, [0, 1], [focus[1], H * 0.5]);
  const x = sx - focus[0] * k;
  const y = sy - focus[1] * k;
  // Entrada tipo dron: muy inclinado al principio, se va enderezando
  const tilt = interpolate(
    frame,
    [T.titleOut, T.routeEnd, T.zoomEnd],
    [58, 38, 30],
    {
      ...clamp,
      easing: Easing.out(Easing.cubic),
    },
  );
  const rotZ = interpolate(
    frame,
    [T.titleOut, T.zoomEnd, T.closingIn],
    [-10, -2, 2],
    clamp,
  );
  const lift = interpolate(
    frame,
    [T.titleOut, T.routeStart + 20],
    [1.35, 1.08],
    {
      ...clamp,
      easing: Easing.out(Easing.cubic),
    },
  );
  return {
    k,
    x,
    y,
    tilt,
    rotZ,
    lift,
    toPlane: (pt: [number, number]): [number, number] => [
      pt[0] * k + x,
      pt[1] * k + y,
    ],
  };
};

// Marca en el plano con etiqueta "de pie" mirando a la cámara
const Marker: React.FC<{
  at: [number, number];
  show: number;
  tilt: number;
  color?: string;
  label?: string;
  sub?: string;
  height?: number;
  dx?: number;
}> = ({ at, show, tilt, color = ACCENT, label, sub, height = 120, dx = 0 }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  if (frame < show) {
    return null;
  }
  const s = spring({ frame: frame - show, fps, config: { damping: 14 } });
  const pole = spring({
    frame: frame - show - 4,
    fps,
    config: { damping: 20 },
  });
  const pulse = ((frame - show) / 36) % 1;
  return (
    <div
      style={{
        position: "absolute",
        left: at[0],
        top: at[1],
        transformStyle: "preserve-3d",
      }}
    >
      {/* Onda en el suelo */}
      <div
        style={{
          position: "absolute",
          width: 120,
          height: 120,
          left: -60,
          top: -60,
          borderRadius: "50%",
          border: `3px solid ${color}`,
          opacity: (1 - pulse) * 0.8,
          transform: `scale(${0.2 + pulse})`,
        }}
      />
      <div
        style={{
          position: "absolute",
          width: 18,
          height: 18,
          left: -9,
          top: -9,
          borderRadius: "50%",
          background: color,
          boxShadow: `0 0 24px 6px ${color}`,
          transform: `scale(${s})`,
        }}
      />
      {label ? (
        // Panel vertical: se contrarresta la inclinación del plano
        <div
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            transformOrigin: "0 0",
            transform: `rotateX(${-tilt}deg)`,
            transformStyle: "preserve-3d",
          }}
        >
          <div
            style={{
              position: "absolute",
              left: -1,
              bottom: 0,
              width: 2,
              height: height * pole,
              background: `linear-gradient(to top, ${color}, rgba(255,255,255,0.0))`,
            }}
          />
          <div
            style={{
              position: "absolute",
              left: dx,
              bottom: height * pole,
              transform: `translateX(-50%) translateY(${(1 - s) * 20}px)`,
              opacity: Math.min(1, s),
              whiteSpace: "nowrap",
              padding: "14px 22px",
              borderRadius: 14,
              background: "rgba(8,16,32,0.72)",
              border: `1px solid ${color}55`,
              boxShadow: `0 0 40px ${color}33`,
              backdropFilter: "blur(8px)",
              textAlign: "center",
            }}
          >
            <div
              style={{
                fontFamily: SANS,
                fontWeight: 800,
                fontSize: 30,
                color: WHITE,
                letterSpacing: "-0.01em",
              }}
            >
              {label}
            </div>
            {sub ? (
              <div
                style={{ fontFamily: MONO, fontSize: 17, color, marginTop: 4 }}
              >
                {sub}
              </div>
            ) : null}
          </div>
        </div>
      ) : null}
    </div>
  );
};

const MapPlane: React.FC = () => {
  const frame = useCurrentFrame();
  const cam = useCamera();
  const routeP = interpolate(frame, [T.routeStart, T.routeEnd], [0, 1], {
    ...clamp,
    easing: Easing.inOut(Easing.quad),
  });
  const head = getPointAtLength(routeD, routeLength * routeP);
  const headPlane = head ? cam.toPlane([head.x, head.y]) : [0, 0];
  const trail = routeLength * 0.12;

  const branch = (to: LonLat) =>
    `M${project(GUATE).join(",")} L${project(to).join(",")}`;
  const branchP = interpolate(frame, [T.step3 + 5, T.step3 + 30], [0, 1], {
    ...clamp,
    easing: ease,
  });
  const gtGlow = interpolate(
    frame,
    [T.routeEnd - 10, T.zoomEnd],
    [0, 1],
    clamp,
  );
  const routeDim = interpolate(
    frame,
    [T.zoomEnd - 10, T.zoomEnd + 15],
    [1, 0.35],
    clamp,
  );
  const appear = interpolate(
    frame,
    [T.titleOut - 10, T.titleOut + 25],
    [0, 1],
    clamp,
  );

  // Rejilla tipo HUD que se mueve con el mapa
  const grid = 120 * cam.k;

  return (
    <AbsoluteFill style={{ perspective: 1400, perspectiveOrigin: "50% 30%" }}>
      <AbsoluteFill
        style={{
          transformStyle: "preserve-3d",
          transform: `translateY(60px) rotateX(${cam.tilt}deg) rotateZ(${cam.rotZ}deg) scale(${cam.lift})`,
          opacity: appear,
        }}
      >
        <div
          style={{
            position: "absolute",
            left: -W,
            top: -H,
            width: W * 3,
            height: H * 3,
            backgroundImage: `linear-gradient(${BORDER}33 1px, transparent 1px), linear-gradient(90deg, ${BORDER}33 1px, transparent 1px)`,
            backgroundSize: `${grid}px ${grid}px`,
            backgroundPosition: `${cam.x + W}px ${cam.y + H}px`,
          }}
        />
        <svg
          width={W}
          height={H}
          style={{ position: "absolute", overflow: "visible" }}
        >
          <defs>
            <filter id="glow" x="-50%" y="-50%" width="200%" height="200%">
              <feGaussianBlur stdDeviation={6 / cam.k} result="b" />
              <feMerge>
                <feMergeNode in="b" />
                <feMergeNode in="b" />
                <feMergeNode in="SourceGraphic" />
              </feMerge>
            </filter>
          </defs>
          <g transform={`translate(${cam.x} ${cam.y}) scale(${cam.k})`}>
            {shapes.map((f) => {
              const id = String(f.id);
              const isMx = id === "484";
              const isGt = id === "320";
              if (isGt) {
                return null;
              }
              return (
                <path
                  key={id}
                  d={pathGen(f) ?? undefined}
                  fill={isMx ? LAND : LAND_DIM}
                  stroke={isMx ? "#3A6194" : BORDER}
                  strokeWidth={(isMx ? 1.6 : 1) / cam.k}
                  strokeLinejoin="round"
                />
              );
            })}
            {shapes
              .filter((f) => String(f.id) === "320")
              .map((f) => (
                <g key="gt">
                  <path
                    d={pathGen(f) ?? undefined}
                    fill={`rgba(78,168,255,${0.06 + 0.12 * gtGlow})`}
                  />
                  <path
                    d={pathGen(f) ?? undefined}
                    fill="none"
                    stroke={ACCENT}
                    strokeOpacity={0.4 + 0.6 * gtGlow}
                    strokeWidth={2.5 / cam.k}
                    strokeLinejoin="round"
                    filter={gtGlow > 0 ? "url(#glow)" : undefined}
                  />
                </g>
              ))}
            <path
              d={routeD}
              fill="none"
              stroke={GOLD}
              strokeOpacity={0.25 * routeDim}
              strokeWidth={3 / cam.k}
              strokeDasharray={`${3 / cam.k} ${9 / cam.k}`}
              strokeLinecap="round"
            />
            <path
              d={routeD}
              fill="none"
              stroke={GOLD}
              strokeOpacity={0.85 * routeDim}
              strokeWidth={4 / cam.k}
              strokeLinecap="round"
              strokeDasharray={`${routeLength * routeP} ${routeLength + 1}`}
              filter="url(#glow)"
            />
            {routeP > 0 && routeP < 1 ? (
              <path
                d={routeD}
                fill="none"
                stroke="#FFF3D6"
                strokeWidth={6 / cam.k}
                strokeLinecap="round"
                strokeDasharray={`0 ${Math.max(0, routeLength * routeP - trail)} ${Math.min(trail, routeLength * routeP)} ${routeLength * 2}`}
                filter="url(#glow)"
              />
            ) : null}
            {[XELA, ESCUINTLA].map((to) => {
              const d = branch(to);
              const len = getLength(d);
              return (
                <path
                  key={to.join()}
                  d={d}
                  fill="none"
                  stroke={GOLD}
                  strokeWidth={4 / cam.k}
                  strokeLinecap="round"
                  strokeDasharray={`${len * branchP} ${len + 1}`}
                  filter="url(#glow)"
                />
              );
            })}
          </g>
        </svg>

        {routeP > 0 && routeP < 1 ? (
          <div
            style={{
              position: "absolute",
              left: headPlane[0] - 14,
              top: headPlane[1] - 14,
              width: 28,
              height: 28,
              borderRadius: "50%",
              background: "#FFF6E0",
              boxShadow: `0 0 30px 12px ${GOLD}`,
            }}
          />
        ) : null}

        {frame < T.zoomStart + 5 ? (
          <Marker
            at={cam.toPlane(project(MONTERREY))}
            show={T.routeStart - 18}
            tilt={cam.tilt}
            color={GOLD}
            label="Monterrey"
            sub="25.67°N · 100.31°O"
            height={140}
          />
        ) : null}
        {frame < T.zoomStart + 5 ? (
          <Marker
            at={cam.toPlane(project(GUATE))}
            show={T.routeEnd - 6}
            tilt={cam.tilt}
            color={ACCENT}
            label="Ciudad de Guatemala"
            sub="14.63°N · 90.51°O"
            height={150}
          />
        ) : null}

        {/* Pasos sobre Guatemala */}
        <Marker
          at={cam.toPlane(project(GUATE))}
          show={T.step1 + 8}
          tilt={cam.tilt}
          color={ACCENT}
          label="Almacén + socio local"
          sub="Ciudad de Guatemala"
          height={250}
          dx={-40}
        />
        <Marker
          at={cam.toPlane(project([-90.38, 14.7]))}
          show={T.step2 + 6}
          tilt={cam.tilt}
          color={GOLD}
        />
        <Marker
          at={cam.toPlane(project([-90.42, 14.55]))}
          show={T.step2 + 12}
          tilt={cam.tilt}
          color={GOLD}
          label="2 tiendas piloto"
          sub="¡Ofertas de apertura!"
          height={120}
          dx={230}
        />
        <Marker
          at={cam.toPlane(project(XELA))}
          show={T.step3 + 28}
          tilt={cam.tilt}
          color={GOLD}
          label="Quetzaltenango"
          height={110}
          dx={-120}
        />
        <Marker
          at={cam.toPlane(project(ESCUINTLA))}
          show={T.step3 + 34}
          tilt={cam.tilt}
          color={GOLD}
          label="Escuintla"
          height={30}
          dx={250}
        />
      </AbsoluteFill>
      {/* Niebla en el horizonte */}
      <AbsoluteFill
        style={{
          background: `linear-gradient(180deg, ${BG} 0%, ${BG}EE 14%, ${BG}00 42%)`,
        }}
      />
    </AbsoluteFill>
  );
};

// Palabras que suben desde una máscara, una tras otra
const Reveal: React.FC<{
  text: string;
  start: number;
  size: number;
  weight?: number;
  color?: string;
  highlight?: string[];
  stagger?: number;
}> = ({
  text,
  start,
  size,
  weight = 800,
  color = WHITE,
  highlight = [],
  stagger = 4,
}) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return (
    <div
      style={{
        fontFamily: SANS,
        fontWeight: weight,
        fontSize: size,
        lineHeight: 1.04,
        letterSpacing: "-0.035em",
        color,
      }}
    >
      {text.split(" ").map((word, i) => {
        const s = spring({
          frame: frame - start - i * stagger,
          fps,
          config: { damping: 200 },
          durationInFrames: 22,
        });
        const clean = word.replace(/[¿?,.]/g, "");
        return (
          <span
            key={i}
            style={{
              display: "inline-block",
              overflow: "hidden",
              verticalAlign: "top",
              paddingBottom: "0.08em",
              marginRight: "0.24em",
            }}
          >
            <span
              style={{
                display: "inline-block",
                transform: `translateY(${(1 - s) * 110}%)`,
                color: highlight.includes(clean) ? GOLD : undefined,
              }}
            >
              {word}
            </span>
          </span>
        );
      })}
    </div>
  );
};

const Label: React.FC<{
  children: React.ReactNode;
  opacity?: number;
  color?: string;
}> = ({ children, opacity = 1, color = ACCENT }) => (
  <div
    style={{
      fontFamily: MONO,
      fontSize: 22,
      letterSpacing: "0.18em",
      textTransform: "uppercase",
      color,
      opacity,
      display: "flex",
      alignItems: "center",
      gap: 16,
    }}
  >
    <span style={{ width: 40, height: 2, background: color }} />
    {children}
  </div>
);

const Hud: React.FC = () => {
  const frame = useCurrentFrame();
  const a = interpolate(
    frame,
    [T.titleOut + 10, T.titleOut + 30],
    [0, 1],
    clamp,
  );
  const swap = interpolate(
    frame,
    [T.zoomStart, T.zoomStart + 15],
    [0, 1],
    clamp,
  );
  const routeP = interpolate(frame, [T.routeStart, T.routeEnd], [0, 1], {
    ...clamp,
    easing: Easing.inOut(Easing.quad),
  });
  const km = Math.round(routeP * 2100);
  const kmShow = interpolate(
    frame,
    [T.routeStart, T.routeStart + 15],
    [0, 1],
    clamp,
  );
  return (
    <>
      <div
        style={{
          position: "absolute",
          left: 100,
          top: 90,
          opacity: a * (1 - swap),
        }}
      >
        <Label>01 — El reto</Label>
        <div style={{ marginTop: 18 }}>
          <Reveal
            text="De Monterrey a Guatemala"
            start={T.titleOut + 12}
            size={72}
          />
        </div>
      </div>
      <div style={{ position: "absolute", left: 100, top: 90, opacity: swap }}>
        <Label>02 — El plan</Label>
        <div style={{ marginTop: 18 }}>
          <Reveal text="Paso a paso" start={T.zoomStart + 4} size={72} />
        </div>
      </div>
      <div
        style={{
          position: "absolute",
          right: 100,
          top: 90,
          textAlign: "right",
          opacity: kmShow * (1 - swap),
        }}
      >
        <div
          style={{
            fontFamily: MONO,
            fontSize: 20,
            letterSpacing: "0.18em",
            color: MUTED,
          }}
        >
          DISTANCIA POR CARRETERA
        </div>
        <div
          style={{
            fontFamily: SANS,
            fontWeight: 800,
            fontSize: 120,
            letterSpacing: "-0.04em",
            color: WHITE,
            lineHeight: 1,
            fontVariantNumeric: "tabular-nums",
          }}
        >
          {km >= 2100 ? "+" : ""}
          {km.toLocaleString("en-US")}
          <span style={{ color: GOLD, fontSize: 56, marginLeft: 12 }}>km</span>
        </div>
        <div
          style={{ fontFamily: SANS, fontSize: 24, color: MUTED, marginTop: 8 }}
        >
          Hoy todo sale de Nuevo León
        </div>
      </div>
    </>
  );
};

const STEPS = [
  {
    at: T.step1,
    n: "01",
    when: "Meses 0–6",
    what: "Socio guatemalteco, almacén y proveedores locales",
  },
  {
    at: T.step2,
    n: "02",
    when: "Meses 6–24",
    what: "2 tiendas de cercanía en la capital, y medir",
  },
  {
    at: T.step3,
    n: "03",
    when: "Mes 24",
    what: "Si cumple metas: Quetzaltenango y Escuintla",
  },
];

const Steps: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const out = interpolate(
    frame,
    [T.closingIn - 5, T.closingIn + 10],
    [1, 0],
    clamp,
  );
  return (
    <div
      style={{
        position: "absolute",
        left: 100,
        top: 300,
        width: 560,
        opacity: out,
      }}
    >
      {STEPS.map((s, i) => {
        const p = spring({
          frame: frame - s.at,
          fps,
          config: { damping: 200 },
          durationInFrames: 20,
        });
        const next = STEPS[i + 1]?.at ?? Infinity;
        const active = frame >= s.at && frame < next;
        const bar = interpolate(
          frame,
          [s.at, Math.min(next, T.closingIn)],
          [0, 1],
          clamp,
        );
        return (
          <div
            key={s.n}
            style={{
              position: "relative",
              padding: "26px 0 26px 34px",
              borderLeft: `2px solid ${active ? GOLD : "rgba(141,162,192,0.25)"}`,
              opacity: p * (active ? 1 : 0.5),
              transform: `translateX(${(1 - p) * -40}px)`,
            }}
          >
            {active ? (
              <div
                style={{
                  position: "absolute",
                  left: -2,
                  top: 0,
                  width: 2,
                  height: `${bar * 100}%`,
                  background: WHITE,
                  boxShadow: `0 0 12px ${GOLD}`,
                }}
              />
            ) : null}
            <div style={{ display: "flex", alignItems: "baseline", gap: 18 }}>
              <span
                style={{
                  fontFamily: MONO,
                  fontWeight: 500,
                  fontSize: 22,
                  color: active ? GOLD : MUTED,
                }}
              >
                {s.n}
              </span>
              <span
                style={{
                  fontFamily: MONO,
                  fontSize: 22,
                  letterSpacing: "0.12em",
                  textTransform: "uppercase",
                  color: active ? GOLD : MUTED,
                }}
              >
                {s.when}
              </span>
            </div>
            <div
              style={{
                fontFamily: SANS,
                fontWeight: active ? 800 : 600,
                fontSize: active ? 40 : 30,
                lineHeight: 1.15,
                letterSpacing: "-0.02em",
                color: WHITE,
                marginTop: 8,
              }}
            >
              {s.what}
            </div>
          </div>
        );
      })}
    </div>
  );
};

const TitleCard: React.FC = () => {
  const frame = useCurrentFrame();
  const out = interpolate(frame, [T.titleOut - 6, T.titleOut + 14], [0, 1], {
    ...clamp,
    easing: Easing.in(Easing.cubic),
  });
  const line = interpolate(frame, [2, 30], [0, 1], { ...clamp, easing: ease });
  if (out >= 1) {
    return null;
  }
  return (
    <AbsoluteFill
      style={{
        justifyContent: "center",
        paddingLeft: 140,
        opacity: 1 - out,
        transform: `scale(${1 + out * 0.08})`,
        filter: `blur(${out * 12}px)`,
      }}
    >
      <Label opacity={line}>Supermercados Merco · Internacionalización</Label>
      <div style={{ marginTop: 30 }}>
        <Reveal text="¿Cómo llegaría Merco" start={8} size={140} />
        <Reveal
          text="a Guatemala?"
          start={18}
          size={140}
          highlight={["Guatemala"]}
        />
      </div>
      <div
        style={{
          marginTop: 34,
          width: 700 * line,
          height: 2,
          background: `linear-gradient(90deg, ${GOLD}, transparent)`,
        }}
      />
    </AbsoluteFill>
  );
};

const ClosingCard: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  if (frame < T.closingIn) {
    return null;
  }
  const veil = interpolate(
    frame,
    [T.closingIn, T.closingIn + 14],
    [0, 1],
    clamp,
  );
  const badge = spring({
    frame: frame - T.closingIn - 26,
    fps,
    config: { damping: 200 },
    durationInFrames: 18,
  });
  return (
    <AbsoluteFill>
      <AbsoluteFill style={{ background: `${BG}`, opacity: veil * 0.86 }} />
      <AbsoluteFill style={{ justifyContent: "center", paddingLeft: 140 }}>
        <Label opacity={veil} color={GOLD}>
          Recomendación final
        </Label>
        <div style={{ marginTop: 26 }}>
          <Reveal
            text="Sí a Guatemala."
            start={T.closingIn + 4}
            size={150}
            highlight={["Guatemala"]}
          />
          <Reveal
            text="Paso a paso, con un socio local."
            start={T.closingIn + 12}
            size={72}
            weight={300}
            stagger={2}
          />
        </div>
        <div
          style={{
            marginTop: 40,
            display: "flex",
            gap: 16,
            opacity: badge,
            transform: `translateY(${(1 - badge) * 16}px)`,
          }}
        >
          {["2027 · Socio y almacén", "2028 · 2 tiendas", "2029 · Crecer"].map(
            (t) => (
              <div
                key={t}
                style={{
                  fontFamily: MONO,
                  fontSize: 22,
                  color: WHITE,
                  padding: "12px 22px",
                  borderRadius: 999,
                  border: "1px solid rgba(255,255,255,0.25)",
                  background: "rgba(255,255,255,0.06)",
                }}
              >
                {t}
              </div>
            ),
          )}
          <div
            style={{
              fontFamily: MONO,
              fontSize: 22,
              color: BG,
              padding: "12px 22px",
              borderRadius: 999,
              background: GOLD,
              fontWeight: 500,
            }}
          >
            Aprobado con condiciones
          </div>
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};

// Partículas de polvo flotando delante de la cámara
const Particles: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{ pointerEvents: "none" }}>
      {new Array(40).fill(0).map((_, i) => {
        const x = random(`x${i}`) * W;
        const y =
          (random(`y${i}`) * H -
            frame * (0.3 + random(`v${i}`) * 0.8) +
            H * 2) %
          H;
        const r = 1 + random(`r${i}`) * 2.5;
        return (
          <div
            key={i}
            style={{
              position: "absolute",
              left: x,
              top: y,
              width: r * 2,
              height: r * 2,
              borderRadius: "50%",
              background: WHITE,
              opacity: 0.08 + random(`o${i}`) * 0.25,
              filter: "blur(1px)",
            }}
          />
        );
      })}
    </AbsoluteFill>
  );
};

const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill
      style={{ pointerEvents: "none", mixBlendMode: "overlay", opacity: 0.35 }}
    >
      <svg width={W} height={H}>
        <filter id="grain">
          <feTurbulence
            type="fractalNoise"
            baseFrequency={0.85}
            numOctaves={2}
            seed={frame % 12}
          />
          <feColorMatrix type="saturate" values="0" />
        </filter>
        <rect width={W} height={H} filter="url(#grain)" />
      </svg>
    </AbsoluteFill>
  );
};

export const MercoGuatemala: React.FC = () => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const stepsVisible = frame >= T.step1 - 5 && frame < T.closingIn + 20;
  const fadeIn = interpolate(frame, [0, 8], [0, 1], clamp);
  const fadeOut = interpolate(
    frame,
    [durationInFrames - 10, durationInFrames],
    [1, 0],
    clamp,
  );
  return (
    <AbsoluteFill
      style={{
        background: `radial-gradient(ellipse at 60% 40%, #0C1A33 0%, ${BG} 65%)`,
      }}
    >
      <AbsoluteFill style={{ opacity: fadeIn * fadeOut }}>
        <MapPlane />
        <Particles />
        <Hud />
        {stepsVisible ? <Steps /> : null}
        <TitleCard />
        <ClosingCard />
        {/* Viñeta */}
        <AbsoluteFill
          style={{
            background:
              "radial-gradient(ellipse at center, transparent 55%, rgba(0,0,0,0.65) 100%)",
          }}
        />
        <Grain />
      </AbsoluteFill>
      <Audio
        src={staticFile("musica-pitch.mp3")}
        volume={(f) =>
          interpolate(f, [durationInFrames - 25, durationInFrames], [0.9, 0], {
            extrapolateLeft: "clamp",
          })
        }
      />
    </AbsoluteFill>
  );
};
