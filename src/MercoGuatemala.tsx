import "@fontsource/oswald/700.css";
import "@fontsource/inter/400.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
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
  spring,
  staticFile,
  useCurrentFrame,
  useVideoConfig,
} from "remotion";

// Colores de la presentación
const NAVY = "#0F2D5C";
const BLUE = "#2B6CB0";
const LIGHT = "#E6EEF8";
const OCEAN = "#D4E2F2";
const WHITE = "#FFFFFF";
const HEAD = "'Oswald', sans-serif";
const BODY = "'Inter', sans-serif";

const W = 1920;
const H = 1080;

// Momentos clave (en fotogramas, 30 fps)
const T = {
  titleOut: 80,
  routeStart: 105,
  routeEnd: 200,
  zoomStart: 215,
  zoomEnd: 255,
  step1: 255,
  step2: 300,
  step3: 345,
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
        [-121, 33.5],
        [-80, 11.5],
      ],
    },
  },
);
const pathGen = geoPath(projection);
const project = (p: LonLat) => projection(p) as [number, number];

const routeD = line().curve(curveCatmullRom.alpha(0.5))(ROUTE.map(project))!;
const routeLength = getLength(routeD);

// Encuadre final: Guatemala a la derecha, hueco a la izquierda para los pasos
const ZOOM = 3.8;
const guateCenter = project([-90.5, 15.2]);

const useCamera = () => {
  const frame = useCurrentFrame();
  const p = interpolate(frame, [T.zoomStart, T.zoomEnd], [0, 1], {
    ...clamp,
    easing: ease,
  });
  // Escala en logarítmico para que el zoom se sienta uniforme; el centro de
  // Guatemala viaja de su posición inicial a su sitio final en pantalla.
  const k = Math.exp(interpolate(p, [0, 1], [0, Math.log(ZOOM)]));
  const sx = interpolate(p, [0, 1], [guateCenter[0], W * 0.68]);
  const sy = interpolate(p, [0, 1], [guateCenter[1], H * 0.5]);
  const x = sx - guateCenter[0] * k;
  const y = sy - guateCenter[1] * k;
  return {
    k,
    x,
    y,
    toScreen: (pt: [number, number]): [number, number] => [
      pt[0] * k + x,
      pt[1] * k + y,
    ],
  };
};

const Pin: React.FC<{
  at: [number, number];
  show: number;
  color?: string;
  size?: number;
  label?: string;
  sub?: string;
  side?: "left" | "right";
  // Si se indica, la etiqueta se coloca separada y unida con una línea guía
  offset?: [number, number];
}> = ({
  at,
  show,
  color = BLUE,
  size = 26,
  label,
  sub,
  side = "right",
  offset,
}) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const s = spring({ frame: frame - show, fps, config: { damping: 12 } });
  const pulse = (frame - show) / 30;
  if (frame < show) {
    return null;
  }
  return (
    <div
      style={{
        position: "absolute",
        left: at[0],
        top: at[1],
        transform: "translate(-50%, -50%)",
      }}
    >
      <div
        style={{
          position: "absolute",
          left: "50%",
          top: "50%",
          width: size * 3,
          height: size * 3,
          marginLeft: (-size * 3) / 2,
          marginTop: (-size * 3) / 2,
          borderRadius: "50%",
          border: `3px solid ${color}`,
          opacity: Math.max(0, 0.7 - (pulse % 1) * 0.7),
          transform: `scale(${0.4 + (pulse % 1) * 0.8})`,
        }}
      />
      <div
        style={{
          width: size,
          height: size,
          borderRadius: "50%",
          background: color,
          border: `${size * 0.18}px solid ${WHITE}`,
          boxShadow: "0 4px 14px rgba(15,45,92,0.35)",
          transform: `scale(${s})`,
        }}
      />
      {label && offset ? (
        <svg
          style={{
            position: "absolute",
            left: size / 2,
            top: size / 2,
            overflow: "visible",
          }}
          width={1}
          height={1}
        >
          <line
            x1={0}
            y1={0}
            x2={offset[0] * Math.min(1, s)}
            y2={offset[1] * Math.min(1, s)}
            stroke={NAVY}
            strokeWidth={3}
          />
        </svg>
      ) : null}
      {label ? (
        <div
          style={{
            position: "absolute",
            ...(offset
              ? {
                  left: size / 2 + offset[0],
                  top: size / 2 + offset[1],
                  transform: `translate(-50%, -50%) scale(${0.6 + 0.4 * Math.min(1, s)})`,
                }
              : {
                  top: "50%",
                  [side === "right" ? "left" : "right"]: size + 10,
                  transform: `translateY(-50%) translateX(${(1 - s) * (side === "right" ? -20 : 20)}px)`,
                }),
            opacity: Math.min(1, s),
            whiteSpace: "nowrap",
            textAlign: side === "right" ? "left" : "right",
            background: WHITE,
            padding: "10px 18px",
            borderRadius: 12,
            boxShadow: "0 6px 20px rgba(15,45,92,0.18)",
          }}
        >
          <div
            style={{
              fontFamily: HEAD,
              fontSize: 30,
              color: NAVY,
              textTransform: "uppercase",
              letterSpacing: "0.02em",
            }}
          >
            {label}
          </div>
          {sub ? (
            <div style={{ fontFamily: BODY, fontSize: 20, color: "#4A5D7A" }}>
              {sub}
            </div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
};

const MapLayer: React.FC = () => {
  const frame = useCurrentFrame();
  const cam = useCamera();
  const routeP = interpolate(frame, [T.routeStart, T.routeEnd], [0, 1], {
    ...clamp,
    easing: Easing.inOut(Easing.quad),
  });
  const truck = getPointAtLength(routeD, routeLength * routeP);
  const truckScreen = truck ? cam.toScreen([truck.x, truck.y]) : [0, 0];
  const km = Math.round(routeP * 2100);

  // Rutas de expansión dentro de Guatemala (paso 3)
  const branch = (to: LonLat) =>
    `M${project(GUATE).join(",")} L${project(to).join(",")}`;
  const branchP = interpolate(frame, [T.step3 + 5, T.step3 + 30], [0, 1], {
    ...clamp,
    easing: ease,
  });

  const guateHighlight = interpolate(
    frame,
    [T.zoomStart, T.zoomEnd],
    [0, 1],
    clamp,
  );
  const routeFade = interpolate(
    frame,
    [T.zoomEnd - 10, T.zoomEnd + 15],
    [1, 0.25],
    clamp,
  );

  return (
    <AbsoluteFill style={{ background: OCEAN }}>
      <svg width={W} height={H} style={{ position: "absolute" }}>
        <g transform={`translate(${cam.x} ${cam.y}) scale(${cam.k})`}>
          {shapes.map((f) => {
            const id = String(f.id);
            const isMx = id === "484";
            const isGt = id === "320";
            return (
              <path
                key={id}
                d={pathGen(f) ?? undefined}
                fill={
                  isGt
                    ? interpolateColor(guateHighlight, "#BFD5EE", BLUE)
                    : isMx
                      ? WHITE
                      : "#EEF2F7"
                }
                stroke={isGt ? NAVY : "#9DB3CF"}
                strokeWidth={(isGt || isMx ? 2 : 1) / cam.k}
                strokeLinejoin="round"
              />
            );
          })}
          {/* Ruta completa en punteado y progreso en sólido */}
          <path
            d={routeD}
            fill="none"
            stroke={NAVY}
            strokeOpacity={0.35 * routeFade}
            strokeWidth={4 / cam.k}
            strokeDasharray={`${10 / cam.k} ${10 / cam.k}`}
            strokeLinecap="round"
          />
          <path
            d={routeD}
            fill="none"
            stroke={NAVY}
            strokeOpacity={routeFade}
            strokeWidth={7 / cam.k}
            strokeLinecap="round"
            strokeDasharray={`${routeLength * routeP} ${routeLength + 1}`}
          />
          {[XELA, ESCUINTLA].map((to) => {
            const d = branch(to);
            const len = getLength(d);
            return (
              <path
                key={to.join()}
                d={d}
                fill="none"
                stroke={WHITE}
                strokeWidth={5 / cam.k}
                strokeLinecap="round"
                strokeDasharray={`${len * branchP} ${len + 1}`}
              />
            );
          })}
        </g>
      </svg>

      {/* Etiquetas de la vista general */}
      {frame < T.zoomStart ? (
        <Pin
          at={cam.toScreen(project(MONTERREY))}
          show={T.routeStart - 20}
          color={NAVY}
          label="Monterrey"
          sub="Hoy todo sale de Nuevo León"
        />
      ) : null}
      {frame < T.zoomStart + 10 ? (
        <Pin
          at={cam.toScreen(project(GUATE))}
          show={T.routeEnd - 5}
          color={BLUE}
          label="Ciudad de Guatemala"
          side="right"
        />
      ) : null}

      {/* Camión / punto que recorre la ruta */}
      {routeP > 0 && routeP < 1 ? (
        <div
          style={{
            position: "absolute",
            left: truckScreen[0],
            top: truckScreen[1],
            width: 30,
            height: 30,
            marginLeft: -15,
            marginTop: -15,
            borderRadius: "50%",
            background: "#F2A93B",
            border: `5px solid ${WHITE}`,
            boxShadow: "0 0 0 8px rgba(242,169,59,0.3)",
          }}
        />
      ) : null}

      {/* Contador de distancia */}
      <DistanceBadge km={km} />

      {/* Pines de los pasos (vista de Guatemala) */}
      <Pin
        at={cam.toScreen(project(GUATE))}
        show={T.step1 + 6}
        color={NAVY}
        size={30}
        label="Almacén + socio local"
        sub="Ciudad de Guatemala"
        offset={[-40, -230]}
      />
      <Pin
        at={cam.toScreen(project([-90.4, 14.7]))}
        show={T.step2 + 6}
        color="#F2A93B"
      />
      <Pin
        at={cam.toScreen(project([-90.44, 14.56]))}
        show={T.step2 + 12}
        color="#F2A93B"
        label="2 tiendas piloto"
        sub="¡Ofertas de apertura!"
        offset={[270, -60]}
      />
      <Pin
        at={cam.toScreen(project(XELA))}
        show={T.step3 + 28}
        color={WHITE}
        label="Quetzaltenango"
        offset={[-130, 150]}
      />
      <Pin
        at={cam.toScreen(project(ESCUINTLA))}
        show={T.step3 + 34}
        color={WHITE}
        label="Escuintla"
        offset={[200, 110]}
      />
    </AbsoluteFill>
  );
};

const interpolateColor = (t: number, a: string, b: string) => {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  const c = pa.map((v, i) => Math.round(v + (pb[i] - v) * t));
  return `rgb(${c.join(",")})`;
};

const DistanceBadge: React.FC<{ km: number }> = ({ km }) => {
  const frame = useCurrentFrame();
  const show = interpolate(
    frame,
    [T.routeStart, T.routeStart + 15],
    [0, 1],
    clamp,
  );
  const hide = interpolate(
    frame,
    [T.zoomStart, T.zoomStart + 12],
    [0, 1],
    clamp,
  );
  return (
    <div
      style={{
        position: "absolute",
        right: 90,
        top: 90,
        opacity: show * (1 - hide),
        transform: `translateY(${(1 - show) * 20}px)`,
        background: NAVY,
        color: WHITE,
        borderRadius: 18,
        padding: "22px 34px",
        boxShadow: "0 10px 30px rgba(15,45,92,0.3)",
      }}
    >
      <div style={{ fontFamily: BODY, fontSize: 22, opacity: 0.8 }}>
        Distancia por carretera
      </div>
      <div style={{ fontFamily: HEAD, fontSize: 72, lineHeight: 1.1 }}>
        {km >= 2100 ? "+" : ""}
        {km.toLocaleString("en-US")} km
      </div>
    </div>
  );
};

const Header: React.FC = () => {
  const frame = useCurrentFrame();
  const a = interpolate(frame, [T.titleOut, T.titleOut + 20], [0, 1], clamp);
  const swap = interpolate(
    frame,
    [T.zoomStart, T.zoomStart + 20],
    [0, 1],
    clamp,
  );
  const style: React.CSSProperties = {
    position: "absolute",
    left: 90,
    top: 80,
    background: "rgba(255,255,255,0.92)",
    borderRadius: 18,
    padding: "22px 34px",
    boxShadow: "0 10px 30px rgba(15,45,92,0.15)",
  };
  return (
    <>
      <div style={{ ...style, opacity: a * (1 - swap) }}>
        <Kicker>El reto · Paso 2</Kicker>
        <Title size={56}>De Monterrey a Guatemala</Title>
      </div>
      <div style={{ ...style, opacity: swap }}>
        <Kicker>Acción estratégica · Paso 4</Kicker>
        <Title size={56}>Cómo llegamos, paso a paso</Title>
      </div>
    </>
  );
};

const Kicker: React.FC<{ children: React.ReactNode; color?: string }> = ({
  children,
  color = BLUE,
}) => (
  <div
    style={{
      fontFamily: HEAD,
      fontSize: 24,
      letterSpacing: "0.12em",
      textTransform: "uppercase",
      color,
    }}
  >
    {children}
  </div>
);

const Title: React.FC<{
  children: React.ReactNode;
  size: number;
  color?: string;
}> = ({ children, size, color = NAVY }) => (
  <div
    style={{
      fontFamily: HEAD,
      fontSize: size,
      lineHeight: 1.05,
      textTransform: "uppercase",
      color,
    }}
  >
    {children}
  </div>
);

const STEPS = [
  {
    at: T.step1,
    n: "1",
    when: "Meses 0 a 6",
    what: "Socio guatemalteco, almacén y proveedores locales",
  },
  {
    at: T.step2,
    n: "2",
    when: "Meses 6 a 24",
    what: "Abrir 2 tiendas de cercanía en la capital y medir",
  },
  {
    at: T.step3,
    n: "3",
    when: "Mes 24",
    what: "Si cumple las metas, crecer a Quetzaltenango y Escuintla",
  },
];

const Steps: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return (
    <div
      style={{
        position: "absolute",
        left: 90,
        top: 270,
        width: 640,
        display: "flex",
        flexDirection: "column",
        gap: 26,
      }}
    >
      {STEPS.map((s) => {
        const p = spring({ frame: frame - s.at, fps, config: { damping: 16 } });
        const active =
          frame >= s.at &&
          (STEPS.find((o) => o.at > s.at)?.at ?? Infinity) > frame;
        return (
          <div
            key={s.n}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 24,
              background: active ? NAVY : WHITE,
              color: active ? WHITE : NAVY,
              borderRadius: 20,
              padding: "24px 30px",
              boxShadow: "0 10px 30px rgba(15,45,92,0.18)",
              opacity: Math.min(1, p),
              transform: `translateX(${(1 - p) * -60}px)`,
            }}
          >
            <div
              style={{
                flex: "0 0 auto",
                width: 76,
                height: 76,
                borderRadius: "50%",
                background: active ? WHITE : BLUE,
                color: active ? NAVY : WHITE,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: HEAD,
                fontSize: 44,
              }}
            >
              {s.n}
            </div>
            <div>
              <div
                style={{
                  fontFamily: HEAD,
                  fontSize: 34,
                  textTransform: "uppercase",
                  color: active ? "#9EC5F0" : BLUE,
                }}
              >
                {s.when}
              </div>
              <div
                style={{
                  fontFamily: BODY,
                  fontWeight: 600,
                  fontSize: 26,
                  lineHeight: 1.3,
                }}
              >
                {s.what}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
};

const TitleCard: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const out = interpolate(frame, [T.titleOut, T.titleOut + 22], [0, 1], {
    ...clamp,
    easing: Easing.in(Easing.cubic),
  });
  const s1 = spring({ frame: frame - 4, fps, config: { damping: 18 } });
  const s2 = spring({ frame: frame - 14, fps, config: { damping: 18 } });
  const s3 = spring({ frame: frame - 26, fps, config: { damping: 18 } });
  if (out >= 1) {
    return null;
  }
  return (
    <AbsoluteFill
      style={{
        background: NAVY,
        transform: `translateY(${-out * 100}%)`,
        justifyContent: "center",
        paddingLeft: 140,
      }}
    >
      <div style={{ opacity: s1, transform: `translateY(${(1 - s1) * 30}px)` }}>
        <Kicker color="#9EC5F0">
          Supermercados Merco · Internacionalización
        </Kicker>
      </div>
      <div
        style={{
          opacity: s2,
          transform: `translateY(${(1 - s2) * 30}px)`,
          marginTop: 16,
        }}
      >
        <Title size={130} color={WHITE}>
          ¿Cómo llegaría Merco
        </Title>
        <Title size={130} color={WHITE}>
          a{" "}
          <span
            style={{
              display: "inline-block",
              background: WHITE,
              color: NAVY,
              lineHeight: 1,
              padding: "6px 20px 10px",
              margin: "14px 0 0",
              borderRadius: 8,
            }}
          >
            Guatemala
          </span>
          ?
        </Title>
      </div>
      <div
        style={{
          opacity: s3,
          transform: `translateY(${(1 - s3) * 30}px)`,
          marginTop: 30,
          fontFamily: BODY,
          fontSize: 34,
          color: "#C9DAF0",
        }}
      >
        De Monterrey a Ciudad de Guatemala, en 3 pasos
      </div>
    </AbsoluteFill>
  );
};

const ClosingCard: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const inP = interpolate(frame, [T.closingIn, T.closingIn + 18], [0, 1], {
    ...clamp,
    easing: Easing.out(Easing.cubic),
  });
  const s = spring({
    frame: frame - T.closingIn - 10,
    fps,
    config: { damping: 16 },
  });
  const stamp = spring({
    frame: frame - T.closingIn - 24,
    fps,
    config: { damping: 10 },
  });
  if (frame < T.closingIn) {
    return null;
  }
  return (
    <AbsoluteFill
      style={{
        background: NAVY,
        clipPath: `circle(${inP * 120}% at 70% 55%)`,
        justifyContent: "center",
        paddingLeft: 140,
      }}
    >
      <div style={{ opacity: s, transform: `translateY(${(1 - s) * 30}px)` }}>
        <Kicker color="#9EC5F0">Recomendación final</Kicker>
        <div style={{ marginTop: 16 }}>
          <Title size={110} color={WHITE}>
            Sí a Guatemala,
          </Title>
          <Title size={110} color={WHITE}>
            paso a paso, con un
          </Title>
          <Title size={110} color={WHITE}>
            socio local
          </Title>
        </div>
        <div
          style={{
            marginTop: 30,
            fontFamily: BODY,
            fontSize: 32,
            color: "#C9DAF0",
          }}
        >
          2027: socio y almacén · 2028: 2 tiendas · 2029: crecer
        </div>
      </div>
      <div
        style={{
          position: "absolute",
          right: 160,
          top: 200,
          transform: `rotate(-8deg) scale(${stamp})`,
          border: `6px solid ${WHITE}`,
          borderRadius: 14,
          padding: "14px 30px",
          fontFamily: HEAD,
          fontSize: 44,
          color: WHITE,
          textTransform: "uppercase",
          textAlign: "center",
          lineHeight: 1.1,
        }}
      >
        Aprobado con
        <br />
        condiciones
      </div>
    </AbsoluteFill>
  );
};

export const MercoGuatemala: React.FC = () => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const stepsVisible = frame >= T.step1 - 5 && frame < T.closingIn + 20;
  return (
    <AbsoluteFill style={{ background: LIGHT }}>
      <MapLayer />
      <Header />
      {stepsVisible ? <Steps /> : null}
      <TitleCard />
      <ClosingCard />
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
