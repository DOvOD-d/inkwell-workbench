# -*- coding: utf-8 -*-
"""蓝色流体 + 磨砂玻璃竖筋背景。

参考件是一段 WebGL 着色器（React 组件里内嵌的 GLSL）。这里把同一段 GLSL 搬到桌面 GL 上，
但不用 QOpenGLWidget 当背景层：QOpenGLWidget 会被合成到普通控件之上，会把界面盖住。
改成离屏上下文 + FBO 渲染成 QImage，再由本控件用 QPainter 画出来，控件层级完全正常。

两种用法：
  GlassBackdrop       —— 逐帧推进时间的动效版，只在启动页那 1.2 秒用。
  StaticGlassBackdrop —— 主窗口用：着色器只渲一帧，缓存成 QPixmap，绘画只做一次贴图，
                          切页/滚动不再碰 GPU。
拿不到 OpenGL 或着色器失败时，退回 QPainter 画等效的静态径向渐变。
"""
from __future__ import annotations

import struct
import time

from PySide6.QtCore import Qt, QRect, QTimer
from PySide6.QtGui import QColor, QImage, QPainter, QPixmap, QRadialGradient
from PySide6.QtWidgets import QWidget

VERTEX_SHADER = '''
attribute vec2 position;
void main(){ gl_Position = vec4(position, 0., 1.); }
'''

# 与参考实现逐行对应，只做了两处桌面化改动：
# 去掉 GLSL ES 才有的 precision 限定符，并增加 u_dark 一个变量用于黑夜主题。
FRAGMENT_SHADER = '''
uniform vec2 resolution;
uniform float time;
uniform float dark;
float blob(vec2 p, vec2 center, vec2 radius, float angle){
  p -= center; float c = cos(angle), s = sin(angle);
  p = mat2(c, -s, s, c) * p / radius;
  return exp(-dot(p, p));
}
vec3 fluid(vec2 p, float t){
  float s = sin(t), c = cos(t);
  p.x += .025 * sin(p.y * 7.0 + t * 2.0);
  float trunk = blob(p, vec2(.52 + .15*s - .07*sin(t*2.0), .76), vec2(.235, .46), -.12 - .52*s);
  float crown = blob(p, vec2(.77 - .55*s, .20 - .10*sin(t*2.0)), vec2(.235, .23), -.50 + .45*s);
  float baseY = 1.045 - .17*p.x + .075*sin(p.x * 5.0 + t * 2.0);
  float base = exp(-pow((p.y - baseY) / .092, 2.0));
  float hole = blob(p, vec2(.85 - .55*s, .59 - .13*s), vec2(.19 + .025*c, .32), .1 + .5*s);
  float density = (trunk + .56*crown + .90*base) * (1.0 - .78*hole);
  vec3 color = mix(vec3(.957), vec3(.66, .80, .89), smoothstep(.015, .20, density));
  color = mix(color, vec3(.10, .48, .73), smoothstep(.16, .49, density));
  color = mix(color, vec3(.015, .135, .28), smoothstep(.40, .73, density));
  color = mix(color, vec3(.008, .025, .087), smoothstep(.62, .93, density));
  return color;
}
void main(){
  vec2 uv = vec2(gl_FragCoord.x / resolution.x, 1.0 - gl_FragCoord.y / resolution.y);
  // 竖筋是透镜，折射的是流体本身，不是叠一层半透明条纹。
  float rib = fract(uv.x * 71.0) - .5;
  vec2 refracted = uv + vec2(rib * .018, (rib * rib - .0833) * .034);
  float t = time * 6.28318530718 / 28.0;
  vec3 color = fluid(refracted, t) * .40;
  color += fluid(refracted + vec2(0., .025), t) * .24;
  color += fluid(refracted - vec2(0., .025), t) * .24;
  color += fluid(refracted + vec2(0., .055), t) * .06;
  color += fluid(refracted - vec2(0., .055), t) * .06;
  float edge = pow(abs(rib) * 2.0, 10.0);
  float blue = (1.0 - color.r) * smoothstep(.02, .18, color.b - color.r);
  color += vec3(.018, .025, .03) * (rib * .65 + edge * .24) * blue;
  color = mix(color, vec3(.015, .045, .09) + color * .30, dark);
  gl_FragColor = vec4(clamp(color, 0.0, 1.0), 1.0);
}
'''

FRAME_MS = 50          # 20fps：流体本身很慢，看不出差别，省电也省 CPU
RENDER_WIDTH = 640     # 离屏渲染分辨率，再由 QPainter 平滑放大
RENDER_HEIGHT = 400
STATIC_MOMENT = 0.38   # 静态版取的时刻：团块偏右下，左上留白给内容
STATIC_MAX_WIDTH = 1024  # 静态纹理的渲染宽度上限，再大只是浪费
STATIC_DEBOUNCE_MS = 70  # 拖动窗口时攒一攒再重算纹理
VERTICES = (-1.0, -1.0, 1.0, -1.0, -1.0, 1.0, -1.0, 1.0, 1.0, -1.0, 1.0, 1.0)
GL_FLOAT = 0x1406
GL_TRIANGLES = 0x0004
GL_COLOR_BUFFER_BIT = 0x00004000
GL_NO_ERROR = 0


def _enum(owner, *names):
    """PySide6 的枚举在不同版本里带不带层级都不一样，这里挨个试。"""
    for name in names:
        target = owner
        for part in name.split('.'):
            target = getattr(target, part, None)
            if target is None:
                break
        else:
            if target is not None:
                return target
    return None


class _OffscreenRenderer:
    """离屏 GL 上下文 + FBO：把着色器渲成 QImage 交给控件画。"""

    def __init__(self):
        self.ok = False
        self.reason = ''
        self._surface = None
        self._context = None
        self._program = None
        self._buffer = None
        self._vao = None
        self._fbo = None
        self._functions = None
        self._uniforms = ()

    def setup(self):
        try:
            from PySide6.QtGui import QOffscreenSurface, QOpenGLContext
            from PySide6.QtOpenGL import (QOpenGLBuffer, QOpenGLFramebufferObject,
                                          QOpenGLShader, QOpenGLShaderProgram)
            from PySide6.QtOpenGL import QOpenGLFunctions_2_0 as Functions
        except Exception as exc:
            self.reason = f'本机没有可用的 Qt OpenGL 模块：{type(exc).__name__}'
            return False
        surface = QOffscreenSurface()
        surface.create()
        context = QOpenGLContext()
        if not context.create():
            self.reason = '创建 OpenGL 上下文失败（可能没有显卡驱动）'
            return False
        if not context.makeCurrent(surface):
            self.reason = 'OpenGL 上下文无法切换为当前'
            return False
        functions = Functions()
        if not functions.initializeOpenGLFunctions():
            self.reason = '拿不到 OpenGL 函数集'
            context.doneCurrent()
            return False
        vertex = _enum(QOpenGLShader, 'VertexShader', 'Type.VertexShader',
                       'ShaderTypeBit.VertexShader', 'ShaderTypeBit.Vertex', 'Vertex')
        fragment = _enum(QOpenGLShader, 'FragmentShader', 'Type.FragmentShader',
                         'ShaderTypeBit.FragmentShader', 'ShaderTypeBit.Fragment', 'Fragment')
        if vertex is None or fragment is None:
            self.reason = 'PySide6 的着色器类型枚举不认识'
            context.doneCurrent()
            return False
        program = QOpenGLShaderProgram()
        if not program.addShaderFromSourceCode(vertex, VERTEX_SHADER):
            self.reason = '顶点着色器编译失败：' + (program.log() or '')[:120]
            context.doneCurrent()
            return False
        if not program.addShaderFromSourceCode(fragment, FRAGMENT_SHADER):
            self.reason = '片元着色器编译失败：' + (program.log() or '')[:160]
            context.doneCurrent()
            return False
        if not program.link() or not program.bind():
            self.reason = '着色器链接/绑定失败：' + (program.log() or '')[:120]
            context.doneCurrent()
            return False
        try:
            from PySide6.QtOpenGL import QOpenGLVertexArrayObject
            vao = QOpenGLVertexArrayObject()
            vao.create()
            vao.bind()
        except Exception:
            vao = None
        buffer = QOpenGLBuffer(_enum(QOpenGLBuffer, 'VertexBuffer', 'Type.VertexBuffer') or 1)
        if not buffer.create() or not buffer.bind():
            self.reason = '顶点缓冲区创建失败'
            context.doneCurrent()
            return False
        packed = b''.join(struct.pack('<f', value) for value in VERTICES)
        buffer.allocate(packed, len(packed))
        position = program.attributeLocation('position')
        program.enableAttributeArray(position)
        program.setAttributeBuffer(position, GL_FLOAT, 0, 2)
        if int(functions.glGetError()) != GL_NO_ERROR:
            self.reason = '顶点属性设置后 GL 报错'
            context.doneCurrent()
            return False
        self._surface = surface
        self._context = context
        self._program = program
        self._buffer = buffer
        self._vao = vao
        self._functions = functions
        self._uniforms = (program.uniformLocation('time'),
                          program.uniformLocation('resolution'),
                          program.uniformLocation('dark'))
        self._fbo = QOpenGLFramebufferObject(RENDER_WIDTH, RENDER_HEIGHT)
        self.ok = True
        context.doneCurrent()
        return True

    def render(self, moment, dark, width=None, height=None):
        """渲染一帧，返回 QImage；失败返回 None 并由调用方降级。"""
        from PySide6.QtOpenGL import QOpenGLFramebufferObject
        if not self.ok:
            return None
        target_w = int(width or RENDER_WIDTH)
        target_h = int(height or RENDER_HEIGHT)
        if not self._context.makeCurrent(self._surface):
            self.reason = '渲染时无法切换 OpenGL 上下文'
            self.ok = False
            return None
        try:
            if (self._fbo is None or self._fbo.width() != target_w
                    or self._fbo.height() != target_h):
                self._fbo = QOpenGLFramebufferObject(target_w, target_h)
            self._fbo.bind()
            functions = self._functions
            functions.glViewport(0, 0, target_w, target_h)
            functions.glClearColor(0.0, 0.0, 0.0, 1.0)
            functions.glClear(GL_COLOR_BUFFER_BIT)
            self._program.bind()
            moment_loc, size_loc, dark_loc = self._uniforms
            if moment_loc >= 0:
                self._program.setUniformValue1f(moment_loc, float(moment))
            if size_loc >= 0:
                from PySide6.QtGui import QVector2D
                self._program.setUniformValue(size_loc,
                                              QVector2D(float(target_w),
                                                        float(target_h)))
            if dark_loc >= 0:
                self._program.setUniformValue1f(dark_loc, 1.0 if dark else 0.0)
            functions.glDrawArrays(GL_TRIANGLES, 0, 6)
            error = int(functions.glGetError())
            image = self._fbo.toImage()
            self._fbo.release()
            if error != GL_NO_ERROR or image.isNull():
                self.reason = f'渲染报错 GL {error}，已停用动效'
                self.ok = False
                return None
            return image
        finally:
            self._context.doneCurrent()

    def close(self):
        self._fbo = None
        self._program = None
        self._buffer = None
        self._vao = None
        self._context = None
        self._surface = None


class GlassBackdrop(QWidget):
    """一层背景：能动就动，不能动就静态；窗口隐藏或最小化时自动停帧省电。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('glassbackdrop')
        self._dark = False
        self._motion = True
        self._elapsed = 0.0
        self._last = 0.0
        self._image = None
        self._renderer = _OffscreenRenderer()
        self._ready = False
        self._tried = False
        self._timer = QTimer(self)
        self._timer.setInterval(FRAME_MS)
        self._timer.timeout.connect(self._advance)

    # ------------------------------------------------------------------ 对外接口
    def set_dark(self, dark):
        self._dark = bool(dark)
        if self._ready:
            self._paint_frame()
        self.update()

    def set_motion(self, on):
        self._motion = bool(on)
        self._sync_timer()

    def renderer(self):
        return '着色器' if self._ready else '静态渐变'

    def fallback_reason(self):
        return '' if self._ready else (self._renderer.reason or '尚未开始渲染')

    # ------------------------------------------------------------------ 内部
    def _ensure_ready(self):
        if self._ready:
            return True
        if not self._tried:
            self._tried = True
            self._ready = self._renderer.setup()
        return self._ready

    def _paint_frame(self):
        if not self._ensure_ready():
            self._image = None
            return
        image = self._renderer.render(self._elapsed, self._dark)
        if image is None:
            self._image = None
            self._timer.stop()
            return
        self._image = image
        self.update()

    def _advance(self):
        now = time.monotonic()
        delta = 0.0 if not self._last else min(now - self._last, 0.05)
        self._last = now
        self._elapsed += delta
        self._paint_frame()

    def _running(self):
        if not self._motion or not self.isVisible():
            return False
        window = self.window()
        if window is not None and bool(window.windowState() & Qt.WindowState.WindowMinimized):
            return False
        return True

    def _sync_timer(self):
        if self._running():
            if not self._timer.isActive():
                self._timer.start()
        else:
            self._timer.stop()
            self._last = 0.0

    def showEvent(self, event):
        super().showEvent(event)
        self._sync_timer()
        if self._image is None:
            self._paint_frame()

    def hideEvent(self, event):
        super().hideEvent(event)
        self._sync_timer()

    def event(self, received):
        from PySide6.QtCore import QEvent
        if received.type() in (QEvent.Type.ActivationChange, QEvent.Type.WindowActivate,
                               QEvent.Type.WindowDeactivate):
            self._sync_timer()
        return super().event(received)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        width, height = self.width(), self.height()
        if width <= 0 or height <= 0:
            return
        if self._image is not None and not self._image.isNull():
            painter.drawImage(QRect(0, 0, width, height), self._image)
            painter.end()
            return
        paint_static(painter, width, height, self._dark)
        painter.end()


class StaticGlassBackdrop(QWidget):
    """主窗口背景层：着色器只渲一帧存成 QPixmap，只有尺寸或明暗变了才重算。

    动效版每帧都要跑一次 GPU 并把整窗重画，切页时正好赶上界面最忙，会拖出手感。
    静态版平时绘画只有一次贴图，切页/滚动完全不碰 GPU。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('glassbackdrop')
        self._dark = False
        self._texture = None
        self._renderer = _OffscreenRenderer()
        self._tried = False
        self._rebuild = QTimer(self)
        self._rebuild.setSingleShot(True)
        self._rebuild.setInterval(STATIC_DEBOUNCE_MS)
        self._rebuild.timeout.connect(self._refresh)

    # ------------------------------------------------------------------ 对外接口
    def set_dark(self, dark):
        if self._dark == bool(dark):
            return
        self._dark = bool(dark)
        self._refresh()

    def renderer(self):
        return '着色器' if self._renderer.ok else '静态渐变'

    def fallback_reason(self):
        return '' if self._renderer.ok else (self._renderer.reason or '尚未开始渲染')

    def texture(self):
        return self._texture

    # ------------------------------------------------------------------ 内部
    def _ensure_ready(self):
        if not self._tried:
            self._tried = True
            self._renderer.setup()
        return self._renderer.ok

    def _build(self):
        width, height = self.width(), self.height()
        if width <= 0 or height <= 0:
            return
        dpr = max(1.0, float(self.devicePixelRatioF()))
        pixel_w = max(1, int(round(width * dpr)))
        pixel_h = max(1, int(round(height * dpr)))
        shrink = min(1.0, STATIC_MAX_WIDTH / float(max(pixel_w, 1)))
        render_w = max(64, int(round(pixel_w * shrink)))
        render_h = max(64, int(round(pixel_h * shrink)))
        image = self._renderer.render(STATIC_MOMENT, self._dark, render_w, render_h) \
            if self._ensure_ready() else None
        if image is None or image.isNull():
            image = QImage(render_w, render_h, QImage.Format.Format_ARGB32)
            painter = QPainter(image)
            paint_static(painter, render_w, render_h, self._dark)
            painter.end()
        pixmap = QPixmap.fromImage(image).scaled(
            pixel_w, pixel_h, Qt.AspectRatioMode.IgnoreAspectRatio,
            Qt.TransformationMode.SmoothTransformation)
        pixmap.setDevicePixelRatio(dpr)
        self._texture = pixmap

    def _refresh(self):
        self._build()
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        if self._texture is None:
            self._refresh()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._texture is not None:
            self._rebuild.start()    # 拖动过程中先把旧图拉伸顶着，停手再重算

    def paintEvent(self, event):
        if self._texture is None:
            self._build()
        if self._texture is None:
            return
        painter = QPainter(self)
        painter.drawPixmap(QRect(0, 0, self.width(), self.height()), self._texture)
        painter.end()


def paint_static(painter, width, height, dark=False):
    """静态兜底：与着色器同一套颜色端点，深在右下、浅在左上。"""
    if width <= 0 or height <= 0:
        return
    painter.setPen(Qt.PenStyle.NoPen)
    painter.fillRect(0, 0, width, height, QColor('#0b1220' if dark else '#f4f4f4'))
    gradient = QRadialGradient(width * 0.55, height * 0.95, max(width, height) * 0.72)
    stops = (((0.0, '#020820'), (0.45, '#073d69'), (0.70, '#378fbe'), (1.0, '#f4f4f4'))
             if not dark else
             ((0.0, '#01040f'), (0.45, '#052a49'), (0.70, '#1f5d84'), (1.0, '#0b1220')))
    for offset, colour in stops:
        gradient.setColorAt(offset, QColor(colour))
    painter.setBrush(gradient)
    painter.drawRect(0, 0, width, height)
    step = max(6, int(round(width / 71.0)))
    band = max(1, step // 3)
    highlight = QColor(255, 255, 255, 6 if dark else 10)
    x = 0
    while x < width:
        painter.fillRect(x, 0, band, height, highlight)
        x += step
