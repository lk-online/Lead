/*! For license information please see taskpane.js.LICENSE.txt */
!(function () {
  "use strict";
  var t,
    e,
    r,
    n = {
      27091: function (t) {
        t.exports = function (t, e) {
          return (
            e || (e = {}),
            t
              ? ((t = String(t.__esModule ? t.default : t)),
                e.hash && (t += e.hash),
                e.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(t) ? '"'.concat(t, '"') : t)
              : t
          );
        };
      },
      60806: function (t, e, r) {
        t.exports = r.p + "21bdd0f45fd080bdb02d.css";
      },
    },
    o = {};
  function i(t) {
    var e = o[t];
    if (void 0 !== e) return e.exports;
    var r = (o[t] = { exports: {} });
    return n[t](r, r.exports, i), r.exports;
  }
  (i.m = n),
    (i.n = function (t) {
      var e =
        t && t.__esModule
          ? function () {
              return t.default;
            }
          : function () {
              return t;
            };
      return i.d(e, { a: e }), e;
    }),
    (i.d = function (t, e) {
      for (var r in e) i.o(e, r) && !i.o(t, r) && Object.defineProperty(t, r, { enumerable: !0, get: e[r] });
    }),
    (i.g = (function () {
      if ("object" == typeof globalThis) return globalThis;
      try {
        return this || new Function("return this")();
      } catch (t) {
        if ("object" == typeof window) return window;
      }
    })()),
    (i.o = function (t, e) {
      return Object.prototype.hasOwnProperty.call(t, e);
    }),
    (function () {
      var t;
      i.g.importScripts && (t = i.g.location + "");
      var e = i.g.document;
      if (!t && e && (e.currentScript && (t = e.currentScript.src), !t)) {
        var r = e.getElementsByTagName("script");
        if (r.length) for (var n = r.length - 1; n > -1 && !t; ) t = r[n--].src;
      }
      if (!t) throw new Error("Automatic publicPath is not supported in this browser");
      (t = t
        .replace(/#.*$/, "")
        .replace(/\?.*$/, "")
        .replace(/\/[^\/]+$/, "/")),
        (i.p = t);
    })(),
    (i.b = document.baseURI || self.location.href),
    (function () {
      function t(e) {
        return (
          (t =
            "function" == typeof Symbol && "symbol" == typeof Symbol.iterator
              ? function (t) {
                  return typeof t;
                }
              : function (t) {
                  return t && "function" == typeof Symbol && t.constructor === Symbol && t !== Symbol.prototype
                    ? "symbol"
                    : typeof t;
                }),
          t(e)
        );
      }
      function e() {
        e = function () {
          return n;
        };
        var r,
          n = {},
          o = Object.prototype,
          i = o.hasOwnProperty,
          a =
            Object.defineProperty ||
            function (t, e, r) {
              t[e] = r.value;
            },
          c = "function" == typeof Symbol ? Symbol : {},
          u = c.iterator || "@@iterator",
          s = c.asyncIterator || "@@asyncIterator",
          l = c.toStringTag || "@@toStringTag";
        function f(t, e, r) {
          return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e];
        }
        try {
          f({}, "");
        } catch (r) {
          f = function (t, e, r) {
            return (t[e] = r);
          };
        }
        function p(t, e, r, n) {
          var o = e && e.prototype instanceof b ? e : b,
            i = Object.create(o.prototype),
            c = new _(n || []);
          return a(i, "_invoke", { value: j(t, r, c) }), i;
        }
        function h(t, e, r) {
          try {
            return { type: "normal", arg: t.call(e, r) };
          } catch (t) {
            return { type: "throw", arg: t };
          }
        }
        n.wrap = p;
        var d = "suspendedStart",
          y = "suspendedYield",
          m = "executing",
          v = "completed",
          g = {};
        function b() {}
        function w() {}
        function x() {}
        var k = {};
        f(k, u, function () {
          return this;
        });
        var E = Object.getPrototypeOf,
          I = E && E(E(C([])));
        I && I !== o && i.call(I, u) && (k = I);
        var S = (x.prototype = b.prototype = Object.create(k));
        function L(t) {
          ["next", "throw", "return"].forEach(function (e) {
            f(t, e, function (t) {
              return this._invoke(e, t);
            });
          });
        }
        function O(e, r) {
          function n(o, a, c, u) {
            var s = h(e[o], e, a);
            if ("throw" !== s.type) {
              var l = s.arg,
                f = l.value;
              return f && "object" == t(f) && i.call(f, "__await")
                ? r.resolve(f.__await).then(
                    function (t) {
                      n("next", t, c, u);
                    },
                    function (t) {
                      n("throw", t, c, u);
                    }
                  )
                : r.resolve(f).then(
                    function (t) {
                      (l.value = t), c(l);
                    },
                    function (t) {
                      return n("throw", t, c, u);
                    }
                  );
            }
            u(s.arg);
          }
          var o;
          a(this, "_invoke", {
            value: function (t, e) {
              function i() {
                return new r(function (r, o) {
                  n(t, e, r, o);
                });
              }
              return (o = o ? o.then(i, i) : i());
            },
          });
        }
        function j(t, e, n) {
          var o = d;
          return function (i, a) {
            if (o === m) throw new Error("Generator is already running");
            if (o === v) {
              if ("throw" === i) throw a;
              return { value: r, done: !0 };
            }
            for (n.method = i, n.arg = a; ; ) {
              var c = n.delegate;
              if (c) {
                var u = A(c, n);
                if (u) {
                  if (u === g) continue;
                  return u;
                }
              }
              if ("next" === n.method) n.sent = n._sent = n.arg;
              else if ("throw" === n.method) {
                if (o === d) throw ((o = v), n.arg);
                n.dispatchException(n.arg);
              } else "return" === n.method && n.abrupt("return", n.arg);
              o = m;
              var s = h(t, e, n);
              if ("normal" === s.type) {
                if (((o = n.done ? v : y), s.arg === g)) continue;
                return { value: s.arg, done: n.done };
              }
              "throw" === s.type && ((o = v), (n.method = "throw"), (n.arg = s.arg));
            }
          };
        }
        function A(t, e) {
          var n = e.method,
            o = t.iterator[n];
          if (o === r)
            return (
              (e.delegate = null),
              ("throw" === n &&
                t.iterator.return &&
                ((e.method = "return"), (e.arg = r), A(t, e), "throw" === e.method)) ||
                ("return" !== n &&
                  ((e.method = "throw"),
                  (e.arg = new TypeError("The iterator does not provide a '" + n + "' method")))),
              g
            );
          var i = h(o, t.iterator, e.arg);
          if ("throw" === i.type) return (e.method = "throw"), (e.arg = i.arg), (e.delegate = null), g;
          var a = i.arg;
          return a
            ? a.done
              ? ((e[t.resultName] = a.value),
                (e.next = t.nextLoc),
                "return" !== e.method && ((e.method = "next"), (e.arg = r)),
                (e.delegate = null),
                g)
              : a
            : ((e.method = "throw"),
              (e.arg = new TypeError("iterator result is not an object")),
              (e.delegate = null),
              g);
        }
        function T(t) {
          var e = { tryLoc: t[0] };
          1 in t && (e.catchLoc = t[1]),
            2 in t && ((e.finallyLoc = t[2]), (e.afterLoc = t[3])),
            this.tryEntries.push(e);
        }
        function P(t) {
          var e = t.completion || {};
          (e.type = "normal"), delete e.arg, (t.completion = e);
        }
        function _(t) {
          (this.tryEntries = [{ tryLoc: "root" }]), t.forEach(T, this), this.reset(!0);
        }
        function C(e) {
          if (e || "" === e) {
            var n = e[u];
            if (n) return n.call(e);
            if ("function" == typeof e.next) return e;
            if (!isNaN(e.length)) {
              var o = -1,
                a = function t() {
                  for (; ++o < e.length; ) if (i.call(e, o)) return (t.value = e[o]), (t.done = !1), t;
                  return (t.value = r), (t.done = !0), t;
                };
              return (a.next = a);
            }
          }
          throw new TypeError(t(e) + " is not iterable");
        }
        return (
          (w.prototype = x),
          a(S, "constructor", { value: x, configurable: !0 }),
          a(x, "constructor", { value: w, configurable: !0 }),
          (w.displayName = f(x, l, "GeneratorFunction")),
          (n.isGeneratorFunction = function (t) {
            var e = "function" == typeof t && t.constructor;
            return !!e && (e === w || "GeneratorFunction" === (e.displayName || e.name));
          }),
          (n.mark = function (t) {
            return (
              Object.setPrototypeOf ? Object.setPrototypeOf(t, x) : ((t.__proto__ = x), f(t, l, "GeneratorFunction")),
              (t.prototype = Object.create(S)),
              t
            );
          }),
          (n.awrap = function (t) {
            return { __await: t };
          }),
          L(O.prototype),
          f(O.prototype, s, function () {
            return this;
          }),
          (n.AsyncIterator = O),
          (n.async = function (t, e, r, o, i) {
            void 0 === i && (i = Promise);
            var a = new O(p(t, e, r, o), i);
            return n.isGeneratorFunction(e)
              ? a
              : a.next().then(function (t) {
                  return t.done ? t.value : a.next();
                });
          }),
          L(S),
          f(S, l, "Generator"),
          f(S, u, function () {
            return this;
          }),
          f(S, "toString", function () {
            return "[object Generator]";
          }),
          (n.keys = function (t) {
            var e = Object(t),
              r = [];
            for (var n in e) r.push(n);
            return (
              r.reverse(),
              function t() {
                for (; r.length; ) {
                  var n = r.pop();
                  if (n in e) return (t.value = n), (t.done = !1), t;
                }
                return (t.done = !0), t;
              }
            );
          }),
          (n.values = C),
          (_.prototype = {
            constructor: _,
            reset: function (t) {
              if (
                ((this.prev = 0),
                (this.next = 0),
                (this.sent = this._sent = r),
                (this.done = !1),
                (this.delegate = null),
                (this.method = "next"),
                (this.arg = r),
                this.tryEntries.forEach(P),
                !t)
              )
                for (var e in this) "t" === e.charAt(0) && i.call(this, e) && !isNaN(+e.slice(1)) && (this[e] = r);
            },
            stop: function () {
              this.done = !0;
              var t = this.tryEntries[0].completion;
              if ("throw" === t.type) throw t.arg;
              return this.rval;
            },
            dispatchException: function (t) {
              if (this.done) throw t;
              var e = this;
              function n(n, o) {
                return (c.type = "throw"), (c.arg = t), (e.next = n), o && ((e.method = "next"), (e.arg = r)), !!o;
              }
              for (var o = this.tryEntries.length - 1; o >= 0; --o) {
                var a = this.tryEntries[o],
                  c = a.completion;
                if ("root" === a.tryLoc) return n("end");
                if (a.tryLoc <= this.prev) {
                  var u = i.call(a, "catchLoc"),
                    s = i.call(a, "finallyLoc");
                  if (u && s) {
                    if (this.prev < a.catchLoc) return n(a.catchLoc, !0);
                    if (this.prev < a.finallyLoc) return n(a.finallyLoc);
                  } else if (u) {
                    if (this.prev < a.catchLoc) return n(a.catchLoc, !0);
                  } else {
                    if (!s) throw new Error("try statement without catch or finally");
                    if (this.prev < a.finallyLoc) return n(a.finallyLoc);
                  }
                }
              }
            },
            abrupt: function (t, e) {
              for (var r = this.tryEntries.length - 1; r >= 0; --r) {
                var n = this.tryEntries[r];
                if (n.tryLoc <= this.prev && i.call(n, "finallyLoc") && this.prev < n.finallyLoc) {
                  var o = n;
                  break;
                }
              }
              o && ("break" === t || "continue" === t) && o.tryLoc <= e && e <= o.finallyLoc && (o = null);
              var a = o ? o.completion : {};
              return (
                (a.type = t),
                (a.arg = e),
                o ? ((this.method = "next"), (this.next = o.finallyLoc), g) : this.complete(a)
              );
            },
            complete: function (t, e) {
              if ("throw" === t.type) throw t.arg;
              return (
                "break" === t.type || "continue" === t.type
                  ? (this.next = t.arg)
                  : "return" === t.type
                  ? ((this.rval = this.arg = t.arg), (this.method = "return"), (this.next = "end"))
                  : "normal" === t.type && e && (this.next = e),
                g
              );
            },
            finish: function (t) {
              for (var e = this.tryEntries.length - 1; e >= 0; --e) {
                var r = this.tryEntries[e];
                if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), P(r), g;
              }
            },
            catch: function (t) {
              for (var e = this.tryEntries.length - 1; e >= 0; --e) {
                var r = this.tryEntries[e];
                if (r.tryLoc === t) {
                  var n = r.completion;
                  if ("throw" === n.type) {
                    var o = n.arg;
                    P(r);
                  }
                  return o;
                }
              }
              throw new Error("illegal catch attempt");
            },
            delegateYield: function (t, e, n) {
              return (
                (this.delegate = { iterator: C(t), resultName: e, nextLoc: n }),
                "next" === this.method && (this.arg = r),
                g
              );
            },
          }),
          n
        );
      }
      function r(t, e, r, n, o, i, a) {
        try {
          var c = t[i](a),
            u = c.value;
        } catch (t) {
          return void r(t);
        }
        c.done ? e(u) : Promise.resolve(u).then(n, o);
      }
      function n(t) {
        return function () {
          var e = this,
            n = arguments;
          return new Promise(function (o, i) {
            var a = t.apply(e, n);
            function c(t) {
              r(a, o, i, c, u, "next", t);
            }
            function u(t) {
              r(a, o, i, c, u, "throw", t);
            }
            c(void 0);
          });
        };
      }
      function o(t, e) {
        return (
          (function (t) {
            if (Array.isArray(t)) return t;
          })(t) ||
          (function (t, e) {
            var r = null == t ? null : ("undefined" != typeof Symbol && t[Symbol.iterator]) || t["@@iterator"];
            if (null != r) {
              var n,
                o,
                i,
                a,
                c = [],
                u = !0,
                s = !1;
              try {
                if (((i = (r = r.call(t)).next), 0 === e)) {
                  if (Object(r) !== r) return;
                  u = !1;
                } else for (; !(u = (n = i.call(r)).done) && (c.push(n.value), c.length !== e); u = !0);
              } catch (t) {
                (s = !0), (o = t);
              } finally {
                try {
                  if (!u && null != r.return && ((a = r.return()), Object(a) !== a)) return;
                } finally {
                  if (s) throw o;
                }
              }
              return c;
            }
          })(t, e) ||
          i(t, e) ||
          (function () {
            throw new TypeError(
              "Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."
            );
          })()
        );
      }
      function i(t, e) {
        if (t) {
          if ("string" == typeof t) return a(t, e);
          var r = Object.prototype.toString.call(t).slice(8, -1);
          return (
            "Object" === r && t.constructor && (r = t.constructor.name),
            "Map" === r || "Set" === r
              ? Array.from(t)
              : "Arguments" === r || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(r)
              ? a(t, e)
              : void 0
          );
        }
      }
      function a(t, e) {
        (null == e || e > t.length) && (e = t.length);
        for (var r = 0, n = new Array(e); r < e; r++) n[r] = t[r];
        return n;
      }
      function c(t, e) {
        if (!(t instanceof e)) throw new TypeError("Cannot call a class as a function");
      }
      function u(e, r) {
        for (var n = 0; n < r.length; n++) {
          var o = r[n];
          (o.enumerable = o.enumerable || !1),
            (o.configurable = !0),
            "value" in o && (o.writable = !0),
            Object.defineProperty(
              e,
              (void 0,
              (i = (function (e, r) {
                if ("object" !== t(e) || null === e) return e;
                var n = e[Symbol.toPrimitive];
                if (void 0 !== n) {
                  var o = n.call(e, "string");
                  if ("object" !== t(o)) return o;
                  throw new TypeError("@@toPrimitive must return a primitive value.");
                }
                return String(e);
              })(o.key)),
              "symbol" === t(i) ? i : String(i)),
              o
            );
        }
        var i;
      }
      function s(t, e, r) {
        return e && u(t.prototype, e), r && u(t, r), Object.defineProperty(t, "prototype", { writable: !1 }), t;
      }
      var l = {
          offer: {
            cc: "group@ship-around.com",
            intro: "Dear {name},<br>",
            body: "Please find attached:<br>",
            attachments: "{attachments}",
            note: "If you accept our offer, please note that the last page of our quotation is the proforma invoice.<br>",
            closing:
              "We appreciate your interest in Ship-Around for your procurement needs and we are looking forward to your order confirmation.<br>",
            footnote:
              "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
          },
          acknowledge: {
            cc: "group@ship-around.com",
            intro: "Dear {name},<br>",
            body: "Thank you for reaching out to us.<br><br>We have logged your inquiry with reference SALE{lead}.<br>",
            note: "Please include the above reference in any future correspondence.<br>",
            closing: "We appreciate your interest and will get back to you shortly.<br>",
            footnote:
              "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing your request.",
          },
          follow_up: {
            cc: "group@ship-around.com",
            intro: "Dear {name},<br>",
            body: "I am following up regarding our last quotation {quote_reference} for {quote_items}.<br><br>We would like to know if you are still interested in pursuing this order.<br>",
            note: "I have attached said quotation again for your perusal.<br><br>Please let us know of your decision at your earliest convenience and if there is any way we can assist you further.<br>",
            closing: "We appreciate your interest in Ship-Around for your procurement needs.<br>",
            footnote:
              "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
          },
        },
        f = { Q202: "Quotation", DN202: "Delivery Note", PL202: "Packing List", INV202: "Invoice" };
      Office.onReady(function (t) {
        t.host === Office.HostType.Outlook &&
          ((document.getElementById("sideload-msg").style.display = "none"),
          (document.getElementById("app-body").style.display = "flex"),
          (document.getElementById("acknowledge").onclick = d),
          (document.getElementById("prepare-quote-email").onclick = m),
          (document.getElementById("follow-up").onclick = g));
      });
      var p = (function () {
          function t(e) {
            c(this, t), (this.item = e);
          }
          var r, u, p, h;
          return (
            s(t, [
              {
                key: "getEmailContent",
                value: function (t, e) {
                  if (!l[t]) throw new Error("No template found for type: ".concat(t));
                  for (var r = l[t], n = "", i = 0, a = Object.entries(r); i < a.length; i++) {
                    var c = o(a[i], 2),
                      u = c[0],
                      s = c[1];
                    if ("cc" !== u) {
                      for (var f = s, p = 0, h = Object.entries(e); p < h.length; p++) {
                        var d = o(h[p], 2),
                          y = d[0],
                          m = d[1];
                        f = f.replace("{".concat(y, "}"), m);
                      }
                      n += f + "<br>";
                    }
                  }
                  return n;
                },
              },
              {
                key: "addSubject",
                value:
                  ((h = n(
                    e().mark(function t(r) {
                      var n,
                        o = this,
                        i = arguments;
                      return e().wrap(function (t) {
                        for (;;)
                          switch ((t.prev = t.next)) {
                            case 0:
                              return (
                                (n = !(i.length > 1 && void 0 !== i[1]) || i[1]),
                                t.abrupt(
                                  "return",
                                  new Promise(function (t, e) {
                                    o.item.subject.getAsync(function (i) {
                                      var a;
                                      i.status === Office.AsyncResultStatus.Failed
                                        ? e(i.error)
                                        : ((a = n ? r + i.value : r),
                                          o.item.subject.setAsync(a, function (r) {
                                            r.status === Office.AsyncResultStatus.Failed ? e(r.error) : t();
                                          }));
                                    });
                                  })
                                )
                              );
                            case 2:
                            case "end":
                              return t.stop();
                          }
                      }, t);
                    })
                  )),
                  function (t) {
                    return h.apply(this, arguments);
                  }),
              },
              {
                key: "addCC",
                value:
                  ((p = n(
                    e().mark(function t(r) {
                      var n,
                        o = this,
                        c = arguments;
                      return e().wrap(function (t) {
                        for (;;)
                          switch ((t.prev = t.next)) {
                            case 0:
                              return (
                                (n = c.length > 1 && void 0 !== c[1] && c[1]),
                                t.abrupt(
                                  "return",
                                  new Promise(function (t, e) {
                                    o.item.cc.getAsync(function (c) {
                                      if (c.status === Office.AsyncResultStatus.Failed) e(c.error);
                                      else {
                                        var u;
                                        if (n) u = [r];
                                        else {
                                          var s = c.value;
                                          if (s.includes(r)) return void t();
                                          u = [].concat(
                                            (function (t) {
                                              if (Array.isArray(t)) return a(t);
                                            })((l = s)) ||
                                              (function (t) {
                                                if (
                                                  ("undefined" != typeof Symbol && null != t[Symbol.iterator]) ||
                                                  null != t["@@iterator"]
                                                )
                                                  return Array.from(t);
                                              })(l) ||
                                              i(l) ||
                                              (function () {
                                                throw new TypeError(
                                                  "Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."
                                                );
                                              })(),
                                            [r]
                                          );
                                        }
                                        o.item.cc.setAsync(u, function (r) {
                                          r.status === Office.AsyncResultStatus.Failed ? e(r.error) : t();
                                        });
                                      }
                                      var l;
                                    });
                                  })
                                )
                              );
                            case 2:
                            case "end":
                              return t.stop();
                          }
                      }, t);
                    })
                  )),
                  function (t) {
                    return p.apply(this, arguments);
                  }),
              },
              {
                key: "addBody",
                value:
                  ((u = n(
                    e().mark(function t(r) {
                      var n = this;
                      return e().wrap(function (t) {
                        for (;;)
                          switch ((t.prev = t.next)) {
                            case 0:
                              return t.abrupt(
                                "return",
                                new Promise(function (t, e) {
                                  n.item.body.prependAsync(r, { coercionType: Office.CoercionType.Html }, function (r) {
                                    r.status === Office.AsyncResultStatus.Failed ? e(r.error) : t();
                                  });
                                })
                              );
                            case 1:
                            case "end":
                              return t.stop();
                          }
                      }, t);
                    })
                  )),
                  function (t) {
                    return u.apply(this, arguments);
                  }),
              },
              {
                key: "displayErrorInTaskpane",
                value: function (t, e) {
                  var r = document.createElement("div");
                  if (((r.style.color = "red"), (r.textContent = t), e)) {
                    var n = document.getElementById(e);
                    if (n) return (n.style.display = "block"), void n.appendChild(r);
                  }
                  document.body.appendChild(r);
                },
              },
              {
                key: "getDocumentType",
                value: function (t) {
                  for (var e in f) if (t.startsWith(e)) return "".concat(f[e], " ").concat(t);
                  return t;
                },
              },
              {
                key: "generateAttachmentTable",
                value: function (t) {
                  var e = "";
                  return (
                    t &&
                      t.length > 0 &&
                      ((e = '<table style="border-collapse: collapse;">'),
                      t.forEach(function (t, r) {
                        e +=
                          '<tr style="padding: 2px; background-color: #f5f5f5;"><td style="border: 1px solid; padding: 2px 4px;">'
                            .concat(r + 1, '</td><td style="border: 1px solid gray; padding: 2px 4px;">')
                            .concat(t, "</td></tr>");
                      }),
                      (e += "</table>")),
                    e
                  );
                },
              },
              {
                key: "listAttachments",
                value:
                  ((r = n(
                    e().mark(function t() {
                      var r = this;
                      return e().wrap(function (t) {
                        for (;;)
                          switch ((t.prev = t.next)) {
                            case 0:
                              return t.abrupt(
                                "return",
                                new Promise(function (t, e) {
                                  r.item.getAttachmentsAsync(function (n) {
                                    if (n.status === Office.AsyncResultStatus.Succeeded) {
                                      var o = n.value.filter(function (t) {
                                        return (
                                          t.attachmentType === Office.MailboxEnums.AttachmentType.File && !t.isInline
                                        );
                                      });
                                      if (o && o.length > 0) {
                                        var i = o.map(function (t) {
                                          var e = t.name.split(".").slice(0, -1).join(".");
                                          return (e = r.getDocumentType(e)), r.capitalizeFirstLetter(e);
                                        });
                                        t(i);
                                      } else console.log("The current message has no file attachments."), t([]);
                                    } else console.error("Failed to get attachments:", n.error), e(n.error);
                                  });
                                })
                              );
                            case 1:
                            case "end":
                              return t.stop();
                          }
                      }, t);
                    })
                  )),
                  function () {
                    return r.apply(this, arguments);
                  }),
              },
              {
                key: "capitalizeFirstLetter",
                value: function (t) {
                  return t.charAt(0).toUpperCase() + t.slice(1);
                },
              },
            ]),
            t
          );
        })(),
        h = (function () {
          function t(e, r, n, o) {
            c(this, t),
              (this.modal = document.getElementById(e)),
              (this.allInputDivs = Array.from(this.modal.querySelectorAll("div[id$='InputDiv']"))),
              (this.inputDivs = r.map(function (t) {
                return document.getElementById(t);
              })),
              (this.okButton = document.getElementById(n)),
              (this.cancelButton = document.getElementById(o)),
              this.setupEventListeners();
          }
          return (
            s(t, [
              {
                key: "setupEventListeners",
                value: function () {
                  var t = this;
                  this.okButton.disabled = !0;
                  var e = this.inputDivs.map(function (t) {
                    return t.querySelector("input");
                  });
                  e.forEach(function (r) {
                    r.addEventListener("input", function () {
                      var r = e.every(function (t) {
                        return !t.hasAttribute("required") || "" !== t.value.trim();
                      });
                      t.okButton.disabled = !r;
                    });
                  }),
                    (this.okButton.onclick = function () {
                      var r = e.map(function (t) {
                        return t.value;
                      });
                      t.resolve(r), t.clearInputs(), t.hide();
                    }),
                    (this.cancelButton.onclick = function () {
                      t.reject(new Error("User cancelled the input.")), t.clearInputs(), t.hide();
                    });
                },
              },
              {
                key: "clearInputs",
                value: function () {
                  this.inputDivs.forEach(function (t) {
                    var e = t.querySelector("input");
                    e && (e.value = "");
                  });
                },
              },
              {
                key: "show",
                value: function () {
                  var t = this;
                  return (
                    this.allInputDivs.forEach(function (t) {
                      return (t.style.display = "none");
                    }),
                    this.inputDivs.forEach(function (t) {
                      return (t.style.display = "block");
                    }),
                    new Promise(function (e, r) {
                      (t.modal.style.display = "block"),
                        (t.resolve = e),
                        (t.reject = r),
                        setTimeout(function () {
                          var e = t.inputDivs[0].querySelector("input");
                          e && e.focus();
                        }, 100);
                    })
                  );
                },
              },
              {
                key: "hide",
                value: function () {
                  this.modal.style.display = "none";
                },
              },
            ]),
            t
          );
        })();
      function d() {
        return y.apply(this, arguments);
      }
      function y() {
        return (y = n(
          e().mark(function t() {
            var r, n, i, a, c, u, s, f, d;
            return e().wrap(
              function (t) {
                for (;;)
                  switch ((t.prev = t.next)) {
                    case 0:
                      return (
                        (t.prev = 0),
                        (n = Office.context.mailbox.item),
                        (r = new p(n)),
                        (i = new h("inputModal", ["leadInputDiv", "nameInputDiv"], "modalOk", "modalCancel")),
                        (t.next = 6),
                        i.show()
                      );
                    case 6:
                      return (
                        (a = t.sent),
                        (c = o(a, 2)),
                        (u = c[0]),
                        (s = c[1]),
                        (t.next = 12),
                        r.addSubject("[SALE".concat(u.trim(), "] "))
                      );
                    case 12:
                      return (f = l.acknowledge.cc), (t.next = 15), r.addCC(f);
                    case 15:
                      return (
                        (d = r.getEmailContent("acknowledge", { name: s.trim(), lead: u.trim() })),
                        (t.next = 18),
                        r.addBody(d)
                      );
                    case 18:
                      t.next = 23;
                      break;
                    case 20:
                      (t.prev = 20),
                        (t.t0 = t.catch(0)),
                        r.displayErrorInTaskpane("Error in acknowledgeRFQ: ".concat(t.t0.message), "errorLog");
                    case 23:
                    case "end":
                      return t.stop();
                  }
              },
              t,
              null,
              [[0, 20]]
            );
          })
        )).apply(this, arguments);
      }
      function m() {
        return v.apply(this, arguments);
      }
      function v() {
        return (v = n(
          e().mark(function t() {
            var r, n, i, a, c, u, s, f, d, y, m, v;
            return e().wrap(
              function (t) {
                for (;;)
                  switch ((t.prev = t.next)) {
                    case 0:
                      return (
                        (t.prev = 0),
                        (n = Office.context.mailbox.item),
                        (r = new p(n)),
                        (i = new h("inputModal", ["nameInputDiv"], "modalOk", "modalCancel")),
                        (t.next = 6),
                        i.show()
                      );
                    case 6:
                      return (a = t.sent), (c = o(a, 1)), (u = c[0]), (t.next = 11), r.listAttachments();
                    case 11:
                      if (
                        ((s = t.sent),
                        (f = s
                          .filter(function (t) {
                            return t.startsWith("Quotation Q202");
                          })
                          .map(function (t) {
                            return t.replace("Quotation ", "");
                          })),
                        (d = ""),
                        1 === f.length
                          ? (d = "[Quotation ".concat(f[0], "] "))
                          : f.length > 1 && (d = "[Quotations ".concat(f.join(", "), "] ")),
                        !d)
                      ) {
                        t.next = 18;
                        break;
                      }
                      return (t.next = 18), r.addSubject(d);
                    case 18:
                      return (y = l.offer.cc), (t.next = 21), r.addCC(y);
                    case 21:
                      return (
                        (m = r.generateAttachmentTable(s)),
                        (v = r.getEmailContent("offer", { name: u.trim(), attachments: m })),
                        (t.next = 25),
                        r.addBody(v)
                      );
                    case 25:
                      t.next = 30;
                      break;
                    case 27:
                      (t.prev = 27),
                        (t.t0 = t.catch(0)),
                        r.displayErrorInTaskpane("Error in prepareQuoteEmail: ".concat(t.t0.message), "errorLog");
                    case 30:
                    case "end":
                      return t.stop();
                  }
              },
              t,
              null,
              [[0, 27]]
            );
          })
        )).apply(this, arguments);
      }
      function g() {
        return b.apply(this, arguments);
      }
      function b() {
        return (b = n(
          e().mark(function t() {
            var r, n, i, a, c, u, s, f, d, y, m;
            return e().wrap(
              function (t) {
                for (;;)
                  switch ((t.prev = t.next)) {
                    case 0:
                      return (
                        (t.prev = 0),
                        (n = Office.context.mailbox.item),
                        (r = new p(n)),
                        (i = new h(
                          "inputModal",
                          ["leadInputDiv", "nameInputDiv", "quoteInputDiv", "itemsInputDiv"],
                          "modalOk",
                          "modalCancel"
                        )),
                        (t.next = 6),
                        i.show()
                      );
                    case 6:
                      return (
                        (a = t.sent),
                        (c = o(a, 4)),
                        (u = c[0]),
                        (s = c[1]),
                        (f = c[2]),
                        (d = c[3]),
                        (t.next = 14),
                        r.addSubject("[SALE".concat(u.trim(), "] "))
                      );
                    case 14:
                      return (y = l.follow_up.cc), (t.next = 17), r.addCC(y);
                    case 17:
                      return (
                        (m = r.getEmailContent("follow_up", {
                          name: s.trim(),
                          quote_reference: f.toUpperCase().trim(),
                          quote_items: d.toLowerCase().trim(),
                        })),
                        (t.next = 20),
                        r.addBody(m)
                      );
                    case 20:
                      t.next = 25;
                      break;
                    case 22:
                      (t.prev = 22),
                        (t.t0 = t.catch(0)),
                        r.displayErrorInTaskpane("Error in followUp: ".concat(t.t0.message), "errorLog");
                    case 25:
                    case "end":
                      return t.stop();
                  }
              },
              t,
              null,
              [[0, 22]]
            );
          })
        )).apply(this, arguments);
      }
    })(),
    (t = i(27091)),
    (e = i.n(t)),
    (r = new URL(i(60806), i.b)),
    e()(r);
})();
//# sourceMappingURL=taskpane.js.map
