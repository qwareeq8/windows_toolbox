// Shared UI primitives for Virelo.
// All primitives consume tokens from useTokens() so density/radius/theme
// flow through without prop drilling.

import React from "react";
import { createPortal } from "react-dom";
import { useTokens } from "./theme.jsx";

function Toggle({ on, onChange, size = "md", label, disabled = false }) {
  const t = useTokens();
  const W = size === "sm" ? 26 : 32;
  const H = size === "sm" ? 15 : 18;
  const K = H - 4;
  return (
    <button
      type="button"
      role="switch"
      aria-checked={!!on}
      aria-label={label}
      disabled={disabled}
      onClick={(e) => {
        e.stopPropagation();
        onChange(!on);
      }}
      style={{
        border: "none",
        padding: 0,
        cursor: disabled ? "not-allowed" : "pointer",
        background: "transparent",
        flexShrink: 0,
      }}
    >
      <span
        style={{
          display: "block",
          width: W,
          height: H,
          borderRadius: H / 2,
          background: on ? t.accent : t.isDark ? "rgba(255,255,255,0.12)" : "#D6D2CB",
          position: "relative",
          transition: "background .15s",
        }}
      >
        <span
          style={{
            position: "absolute",
            top: 2,
            left: on ? W - K - 2 : 2,
            width: K,
            height: K,
            borderRadius: K / 2,
            background: "#fff",
            transition: "left .18s cubic-bezier(.2,.7,.3,1)",
            boxShadow: "0 1px 2px rgba(0,0,0,0.2)",
          }}
        />
      </span>
    </button>
  );
}

const Button = React.forwardRef(function Button(
  { children, variant = "secondary", size = "md", onClick, icon, kbd, disabled, ...buttonProps },
  ref,
) {
  const t = useTokens();
  const [hover, setHover] = React.useState(false);
  const ariaDisabled =
    buttonProps["aria-disabled"] === true || buttonProps["aria-disabled"] === "true";
  const unavailable = disabled || ariaDisabled;
  const variants = {
    primary: {
      bg: t.accent,
      color: t.accentOn,
      border: t.accent,
      hover: `color-mix(in oklab, ${t.accent} 90%, #000)`,
    },
    secondary: {
      bg: t.surface,
      color: t.text,
      border: t.borderHi,
      hover: t.hover,
    },
    ghost: {
      bg: "transparent",
      color: t.textDim,
      border: "transparent",
      hover: t.hover,
    },
    danger: {
      bg: "transparent",
      color: t.isDark ? "#F08A7C" : "#A63A2D",
      border: t.borderHi,
      hover: "rgba(197,74,58,0.08)",
    },
  };
  const v = variants[variant];
  const H = size === "sm" ? 26 : 32;
  return (
    <button
      ref={ref}
      type="button"
      onClick={(event) => {
        if (ariaDisabled) {
          event.preventDefault();
          return;
        }
        onClick?.(event);
      }}
      disabled={disabled}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        height: H,
        padding: size === "sm" ? "0 10px" : "0 14px",
        background: hover && !unavailable ? v.hover : v.bg,
        color: v.color,
        border: `1px solid ${v.border === "transparent" ? "transparent" : v.border}`,
        borderRadius: t.radius,
        fontSize: size === "sm" ? 12 : 13,
        fontWeight: 500,
        fontFamily: "inherit",
        cursor: unavailable ? "not-allowed" : "pointer",
        opacity: unavailable ? 0.5 : 1,
        transition: "background .12s, border-color .12s",
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
      }}
      {...buttonProps}
    >
      {icon}
      {children}
      {kbd && (
        <span
          style={{
            marginLeft: 4,
            padding: "1px 5px",
            background: t.hover,
            borderRadius: 3,
            fontSize: 10,
            fontFamily: t.mono,
            opacity: 0.7,
          }}
        >
          {kbd}
        </span>
      )}
    </button>
  );
});

function Card({ title, subtitle, children, footer, padding = true }) {
  const t = useTokens();
  return (
    <div
      style={{
        background: t.surface,
        border: `1px solid ${t.border}`,
        borderRadius: t.radius + 2,
        overflow: "hidden",
        boxShadow: t.shadow,
      }}
    >
      {(title || subtitle) && (
        <div
          style={{
            padding: `${t.rowPad + 2}px ${t.cardPad}px`,
            borderBottom: `1px solid ${t.border}`,
            background: t.surface2,
          }}
        >
          {title && (
            <div
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: t.text,
                letterSpacing: -0.1,
              }}
            >
              {title}
            </div>
          )}
          {subtitle && (
            <div style={{ fontSize: 12, color: t.textDim, marginTop: 2 }}>{subtitle}</div>
          )}
        </div>
      )}
      <div style={{ padding: padding ? `4px ${t.cardPad}px` : 0 }}>{children}</div>
      {footer && (
        <div
          style={{
            padding: `${t.rowPad - 2}px ${t.cardPad}px`,
            borderTop: `1px solid ${t.border}`,
            background: t.surface2,
          }}
        >
          {footer}
        </div>
      )}
    </div>
  );
}

function Row({ label, description, children, last }) {
  const t = useTokens();
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 20,
        padding: `${t.rowPad}px 0`,
        borderBottom: last ? "none" : `1px solid ${t.border}`,
      }}
    >
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 13.5, fontWeight: 500, color: t.text }}>{label}</div>
        {description && (
          <div
            style={{
              fontSize: 12.5,
              color: t.textDim,
              marginTop: 2,
              lineHeight: 1.45,
            }}
          >
            {description}
          </div>
        )}
      </div>
      <div style={{ flexShrink: 0 }}>{children}</div>
    </div>
  );
}

function Segmented({ options, value, onChange, mono, label }) {
  const t = useTokens();
  return (
    <div
      role="group"
      aria-label={label}
      style={{
        display: "inline-flex",
        background: t.isDark ? "rgba(255,255,255,0.04)" : "#EDEAE3",
        border: `1px solid ${t.border}`,
        borderRadius: t.radius + 1,
        padding: 2,
      }}
    >
      {options.map((o) => {
        const val = typeof o === "object" ? o.value : o;
        const label = typeof o === "object" ? o.label : o;
        const active = value === val;
        return (
          <button
            type="button"
            key={val}
            onClick={() => onChange(val)}
            aria-pressed={active}
            style={{
              height: 24,
              padding: "0 10px",
              background: active ? t.surface : "transparent",
              border: "none",
              borderRadius: t.radius - 1,
              color: t.text,
              fontSize: 12,
              fontWeight: active ? 600 : 500,
              fontFamily: mono ? t.mono : "inherit",
              letterSpacing: mono ? 0.3 : 0,
              cursor: "pointer",
              boxShadow: active
                ? t.isDark
                  ? "0 1px 0 rgba(0,0,0,0.4)"
                  : "0 1px 1.5px rgba(0,0,0,0.08)"
                : "none",
              transition: "background .12s",
            }}
          >
            {label}
          </button>
        );
      })}
    </div>
  );
}

function stepBtn(t) {
  return {
    width: 26,
    height: "100%",
    border: "none",
    background: "transparent",
    color: t.textDim,
    fontSize: 14,
    cursor: "pointer",
    fontFamily: "inherit",
  };
}

function Stepper({ value, onChange, min = 1, max = 9999, step = 1, suffix, label }) {
  const t = useTokens();
  return (
    <div
      role="group"
      aria-label={label}
      style={{
        display: "inline-flex",
        alignItems: "center",
        background: t.surface,
        border: `1px solid ${t.borderHi}`,
        borderRadius: t.radius,
        height: 28,
        overflow: "hidden",
      }}
    >
      <button
        type="button"
        onClick={() => onChange(Math.max(min, value - step))}
        aria-label={`Decrease ${label || "value"}`}
        disabled={value <= min}
        style={stepBtn(t)}
      >
        {"−"}
      </button>
      <div
        role="status"
        aria-live="polite"
        aria-label={`${label || "Value"}: ${value}${suffix || ""}`}
        style={{
          minWidth: 48,
          padding: "0 8px",
          fontSize: 13,
          fontWeight: 500,
          color: t.text,
          textAlign: "center",
          fontVariantNumeric: "tabular-nums",
          borderLeft: `1px solid ${t.border}`,
          borderRight: `1px solid ${t.border}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          height: "100%",
          gap: 2,
        }}
      >
        {value}
        {suffix && <span style={{ color: t.textMuted, fontSize: 11 }}>{suffix}</span>}
      </div>
      <button
        type="button"
        onClick={() => onChange(Math.min(max, value + step))}
        aria-label={`Increase ${label || "value"}`}
        disabled={value >= max}
        style={stepBtn(t)}
      >
        +
      </button>
    </div>
  );
}

function Slider({ value, onChange, min = 0, max = 100, label, suffix = "" }) {
  const t = useTokens();
  const pct = ((value - min) / (max - min)) * 100;
  const ref = React.useRef(null);
  const start = (e) => {
    const rect = ref.current.getBoundingClientRect();
    const set = (x) =>
      onChange(
        Math.round(min + Math.max(0, Math.min(1, (x - rect.left) / rect.width)) * (max - min)),
      );
    set(e.clientX);
    const mv = (ev) => set(ev.clientX);
    const up = () => {
      document.removeEventListener("pointermove", mv);
      document.removeEventListener("pointerup", up);
    };
    document.addEventListener("pointermove", mv);
    document.addEventListener("pointerup", up);
  };
  const clamp = (v) => Math.max(min, Math.min(max, v));
  const onKeyDown = (e) => {
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      onChange(clamp(value - 1));
    }
    if (e.key === "ArrowRight") {
      e.preventDefault();
      onChange(clamp(value + 1));
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      onChange(clamp(value - 1));
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      onChange(clamp(value + 1));
    }
    if (e.key === "Home") {
      e.preventDefault();
      onChange(min);
    }
    if (e.key === "End") {
      e.preventDefault();
      onChange(max);
    }
  };
  return (
    <div
      ref={ref}
      role="slider"
      aria-valuemin={min}
      aria-valuemax={max}
      aria-valuenow={value}
      aria-valuetext={`${value}${suffix}`}
      aria-label={label}
      tabIndex={0}
      onKeyDown={onKeyDown}
      onPointerDown={start}
      style={{
        position: "relative",
        height: 20,
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        userSelect: "none",
        minWidth: 160,
      }}
    >
      <div
        style={{
          width: "100%",
          height: 4,
          borderRadius: 2,
          background: t.track,
          position: "relative",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            position: "absolute",
            inset: 0,
            width: `${pct}%`,
            background: t.accent,
            borderRadius: 2,
          }}
        />
      </div>
      <div
        style={{
          position: "absolute",
          left: `calc(${pct}% - 8px)`,
          width: 16,
          height: 16,
          borderRadius: 8,
          background: t.isDark ? t.text : "#fff",
          border: `1.5px solid ${t.accent}`,
          boxShadow: "0 1px 2px rgba(0,0,0,0.1)",
        }}
      />
    </div>
  );
}

function Kbd({ children }) {
  const t = useTokens();
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        minWidth: 22,
        height: 20,
        padding: "0 5px",
        background: t.isDark ? "rgba(255,255,255,0.06)" : "#fff",
        border: `1px solid ${t.borderHi}`,
        borderRadius: 4,
        color: t.text,
        fontSize: 10.5,
        fontWeight: 600,
        letterSpacing: 0.3,
        fontFamily: t.mono,
        boxShadow: t.isDark ? "none" : "0 1px 0 rgba(0,0,0,0.04)",
      }}
    >
      {children}
    </span>
  );
}

function Badge({ children, tone = "default" }) {
  const t = useTokens();
  const tones = {
    default: { bg: t.hover, color: t.textDim },
    accent: { bg: t.accentBg, color: t.accent },
    success: {
      bg: t.isDark ? "rgba(109,212,181,0.14)" : "#E7F0EC",
      color: t.isDark ? "#6DD4B5" : "#2E6F5E",
    },
    warn: {
      bg: t.isDark ? "rgba(255,177,60,0.14)" : "#FBF1DD",
      color: t.isDark ? "#FFC77A" : "#8B6518",
    },
  };
  const v = tones[tone];
  return (
    <span
      style={{
        fontSize: 10.5,
        padding: "2px 7px",
        borderRadius: t.radius - 1,
        background: v.bg,
        color: v.color,
        fontWeight: 600,
        letterSpacing: 0.2,
        display: "inline-block",
        textTransform: "uppercase",
      }}
    >
      {children}
    </span>
  );
}

function Modal({
  open,
  onClose,
  children,
  ariaLabel,
  labelledBy,
  describedBy,
  initialFocusRef,
  width = 360,
  align = "center",
  closeOnBackdrop = true,
}) {
  const t = useTokens();
  const dialogRef = React.useRef(null);
  const previousFocusRef = React.useRef(null);
  const onCloseRef = React.useRef(onClose);

  React.useEffect(() => {
    onCloseRef.current = onClose;
  }, [onClose]);

  React.useEffect(() => {
    if (!open) return undefined;

    previousFocusRef.current = document.activeElement;
    const appShell = document.querySelector("[data-app-shell]");
    const previousAriaHidden = appShell?.getAttribute("aria-hidden");
    const wasInert = appShell?.hasAttribute("inert") ?? false;
    if (appShell) {
      appShell.setAttribute("inert", "");
      appShell.setAttribute("aria-hidden", "true");
    }

    const focusTimer = setTimeout(() => {
      const target =
        initialFocusRef?.current ||
        dialogRef.current?.querySelector(
          'button:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])',
        ) ||
        dialogRef.current;
      target?.focus();
    }, 0);

    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCloseRef.current();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;

      const focusable = Array.from(
        dialogRef.current.querySelectorAll(
          'button:not([disabled]), input:not([disabled]), [href], [tabindex]:not([tabindex="-1"])',
        ),
      );
      if (focusable.length === 0) {
        event.preventDefault();
        dialogRef.current.focus();
        return;
      }
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };
    document.addEventListener("keydown", onKeyDown);

    return () => {
      clearTimeout(focusTimer);
      document.removeEventListener("keydown", onKeyDown);
      if (appShell) {
        if (!wasInert) appShell.removeAttribute("inert");
        if (previousAriaHidden === null) appShell.removeAttribute("aria-hidden");
        else appShell.setAttribute("aria-hidden", previousAriaHidden);
      }
      const previousFocus = previousFocusRef.current;
      if (previousFocus?.isConnected) previousFocus.focus();
    };
  }, [initialFocusRef, open]);

  if (!open) return null;

  return createPortal(
    <div
      onClick={(event) => {
        if (closeOnBackdrop && event.target === event.currentTarget) {
          onCloseRef.current();
        }
      }}
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 100,
        background: t.overlay,
        backdropFilter: "blur(2px)",
        display: "flex",
        alignItems: align === "start" ? "flex-start" : "center",
        justifyContent: "center",
        padding: align === "start" ? "80px 16px 16px" : 16,
      }}
    >
      <div
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-label={ariaLabel}
        aria-labelledby={labelledBy}
        aria-describedby={describedBy}
        tabIndex={-1}
        style={{
          width: `min(${width}px, calc(100vw - 32px))`,
          maxHeight: "calc(100vh - 32px)",
          background: t.surface,
          border: `1px solid ${t.borderHi}`,
          borderRadius: t.radius + 4,
          boxShadow: "0 20px 60px rgba(0,0,0,0.25), 0 4px 12px rgba(0,0,0,0.08)",
          overflow: "auto",
        }}
      >
        {children}
      </div>
    </div>,
    document.body,
  );
}

export { Toggle, Button, Card, Row, Segmented, Stepper, Slider, Kbd, Badge, Modal };
