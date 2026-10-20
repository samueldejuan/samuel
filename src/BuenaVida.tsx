import "@fontsource/playfair-display/400-italic.css";
import "@fontsource/montserrat/400.css";
import {
  AbsoluteFill,
  Audio,
  Easing,
  interpolate,
  OffthreadVideo,
  staticFile,
  useCurrentFrame,
  useVideoConfig,
} from "remotion";

// Texto que entra con desenfoque y sale difuminándose.
const AnimatedText: React.FC<{
  children: React.ReactNode;
  from: number;
  to: number;
  style: React.CSSProperties;
  spacing: [number, number];
}> = ({ children, from, to, style, spacing }) => {
  const frame = useCurrentFrame();
  const ease = Easing.out(Easing.cubic);

  const enter = interpolate(frame, [from, from + 24], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: ease,
  });
  const exit = interpolate(frame, [to - 16, to], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.in(Easing.cubic),
  });
  const letterSpacing = interpolate(frame, [from, to], spacing, {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  if (frame < from || frame > to) {
    return null;
  }

  return (
    <div
      style={{
        ...style,
        opacity: enter * (1 - exit),
        filter: `blur(${(1 - enter) * 14 + exit * 10}px)`,
        transform: `translateY(${(1 - enter) * 40 - exit * 20}px)`,
        letterSpacing: `${letterSpacing}em`,
      }}
    >
      {children}
    </div>
  );
};

export const BuenaVida: React.FC = () => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();

  // Zoom lento tipo Ken Burns
  const scale = interpolate(frame, [0, durationInFrames], [1, 1.08]);
  const fadeIn = interpolate(frame, [0, 12], [0, 1], {
    extrapolateRight: "clamp",
  });
  const fadeOut = interpolate(
    frame,
    [durationInFrames - 12, durationInFrames],
    [1, 0],
    { extrapolateLeft: "clamp" },
  );

  const glow =
    "0 0 30px rgba(255, 170, 90, 0.55), 0 4px 18px rgba(20, 30, 60, 0.45)";

  return (
    <AbsoluteFill style={{ backgroundColor: "black" }}>
      <AbsoluteFill style={{ opacity: fadeIn * fadeOut }}>
        <AbsoluteFill style={{ transform: `scale(${scale})` }}>
          <OffthreadVideo
            src={staticFile("playa.mp4")}
            muted
            style={{ width: "100%", height: "100%", objectFit: "cover" }}
          />
        </AbsoluteFill>
        {/* Tono cálido y viñeta suave para que el texto destaque */}
        <AbsoluteFill
          style={{
            background:
              "linear-gradient(180deg, rgba(25,35,80,0.35) 0%, rgba(25,35,80,0) 45%, rgba(255,140,70,0.08) 70%, rgba(0,0,0,0.25) 100%)",
          }}
        />
        <AbsoluteFill
          style={{
            justifyContent: "flex-start",
            alignItems: "center",
            paddingTop: 560,
          }}
        >
          <AnimatedText
            from={10}
            to={100}
            spacing={[0.02, 0.06]}
            style={{
              position: "absolute",
              fontFamily: "'Playfair Display'",
              fontStyle: "italic",
              fontSize: 132,
              color: "#FFF1DC",
              textShadow: glow,
              whiteSpace: "nowrap",
            }}
          >
            la buena vida
          </AnimatedText>
          <AnimatedText
            from={108}
            to={durationInFrames - 4}
            spacing={[0.55, 0.38]}
            style={{
              position: "absolute",
              fontFamily: "Montserrat",
              fontWeight: 400,
              fontSize: 120,
              color: "#FFD9A0",
              textShadow: glow,
              textTransform: "uppercase",
              whiteSpace: "nowrap",
              // Compensa el espacio extra de la última letra
              marginRight: "-0.4em",
            }}
          >
            Samuel
          </AnimatedText>
        </AbsoluteFill>
      </AbsoluteFill>
      <Audio
        src={staticFile("musica.mp3")}
        volume={(f) =>
          interpolate(f, [durationInFrames - 20, durationInFrames], [1, 0], {
            extrapolateLeft: "clamp",
          })
        }
      />
    </AbsoluteFill>
  );
};
