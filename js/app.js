const App = {
  currentBranch: null,
  currentUser: null,
  users: {},
  branchesData: {},

  init() {
    try {
      const raw = localStorage.getItem("users");
      const parsed = raw ? JSON.parse(raw) : {};
      this.users =
        parsed && typeof parsed === "object" && !Array.isArray(parsed)
          ? parsed
          : {};
    } catch (e) {
      this.users = {};
    }
    try {
      const raw = localStorage.getItem("branchesData");
      const parsed = raw ? JSON.parse(raw) : {};
      this.branchesData =
        parsed && typeof parsed === "object" && !Array.isArray(parsed)
          ? parsed
          : {};
    } catch (e) {
      this.branchesData = {};
    }
    Object.keys(this.users).forEach((k) => {
      const u = this.users[k];
      if (
        !u ||
        typeof u !== "object" ||
        typeof u.user !== "string" ||
        typeof u.pass !== "string"
      ) {
        delete this.users[k];
      }
    });
  },
  saveUsers() {
    localStorage.setItem("users", JSON.stringify(this.users));
  },
  saveData() {
    localStorage.setItem("branchesData", JSON.stringify(this.branchesData));
  },
};
App.init();

function showToast(message, type = "info") {
  const container = document.getElementById("toastContainer");
  if (!container) return;
  const icons = {
    success: "fa-circle-check",
    error: "fa-circle-exclamation",
    info: "fa-circle-info",
  };
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("out");
    setTimeout(() => toast.remove(), 400);
  }, 3200);
}

function confirmModal(title, msg) {
  return new Promise((resolve) => {
    const modal = document.getElementById("confirmModal");
    if (!modal) {
      resolve(false);
      return;
    }
    document.getElementById("confirmTitle").innerHTML =
      `<i class="fas fa-triangle-exclamation"></i> ${title}`;
    document.getElementById("confirmMsg").textContent = msg;
    modal.classList.remove("hidden");
    const yes = document.getElementById("confirmYes");
    const no = document.getElementById("confirmNo");
    const done = (val) => {
      modal.classList.add("hidden");
      yes.onclick = null;
      no.onclick = null;
      resolve(val);
    };
    yes.onclick = () => done(true);
    no.onclick = () => done(false);
  });
}

function setLoading(btn, isLoading) {
  if (!btn) return;
  if (isLoading) {
    btn.dataset.orig = btn.innerHTML;
    btn.classList.add("btn-loading");
    btn.disabled = true;
  } else {
    btn.innerHTML = btn.dataset.orig || btn.innerHTML;
    btn.classList.remove("btn-loading");
    btn.disabled = false;
  }
}

function normalize(str) {
  if (str == null) return "";
  if (typeof str === "number") str = String(str);
  return String(str)
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase()
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/[#*\-_.()\[\]{}<>«»"'،,؟?؟:;؛!\/\\|]/g, "")
    .replace(/[^\w\u0600-\u06FF ]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
function toNum(v) {
  const n = parseFloat(String(v).replace(/[^\d.-]/g, ""));
  return isNaN(n) ? 0 : n;
}

window.addEventListener("load", () => {
  setTimeout(() => {
    const splash = document.getElementById("splashScreen");
    if (splash) splash.classList.add("hidden");
    const dash = document.getElementById("dashboardScreen");
    if (dash) dash.classList.remove("hidden");
    const menuToggle = document.getElementById("menuToggle");
    if (menuToggle)
      menuToggle.onclick = () => {
        document.querySelector(".sidebar").classList.toggle("open");
      };
    bindSidebarNav();
    goToPage("home");
  }, 3000);
});

function bindSidebarNav() {
  document.querySelectorAll("#sidebarNav a").forEach((a) => {
    a.onclick = (e) => {
      e.preventDefault();
      const page = a.dataset.page;
      if (page !== "home" && !App.currentUser) {
        Auth.requestLogin(page);
        return;
      }
      if (page !== "home" && App.currentUser) {
        const user = App.currentUser;
        if (user.role === "developer") {
        } else if (user.role === "admin") {
          if (page !== "admin" && page !== "developer") {
            showToast("هذا المستخدم ليس لديه صلاحية على هذا الفرع", "error");
            return;
          }
        } else if (user.role === "branch") {
          if (page !== user.branch) {
            showToast("هذا المستخدم غير مصرح له بالدخول لهذا الفرع", "error");
            return;
          }
        }
      }
      goToPage(page);
    };
  });
}

function goToPage(page) {
  document.querySelectorAll("#sidebarNav a").forEach((a) => {
    a.classList.toggle("active", a.dataset.page === page);
  });
  document.querySelectorAll(".page").forEach((p) => {
    p.classList.add("hidden");
    p.classList.remove("active");
  });
  const sidebar = document.querySelector(".sidebar");
  if (sidebar) sidebar.classList.remove("open");

  const titles = {
    home: "الصفحة الرئيسية",
    smart: "سمارت فارما",
    san: "صن فارما",
    taiba: "طيبة القصاصين",
    dawa: "الدواء فارما",
    rawi: "الراوي فارما",
    admin: "إنشاء حساب جديد",
    developer: "لوحة المطور",
  };
  const titleEl = document.getElementById("pageTitle");
  if (titleEl) titleEl.textContent = titles[page] || "لوحة التحكم";

  const userLabel = document.getElementById("currentUserLabel");
  const logoutBtn = document.getElementById("logoutBtn");
  if (App.currentUser && userLabel && logoutBtn) {
    const roleLabel =
      App.currentUser.role === "developer"
        ? "مطور"
        : App.currentUser.role === "admin"
          ? "أدمن"
          : "فرع";
    const roleIcon =
      App.currentUser.role === "developer"
        ? "fa-code"
        : App.currentUser.role === "admin"
          ? "fa-user-shield"
          : "fa-user";
    userLabel.innerHTML = `<i class="fas ${roleIcon}"></i> ${App.currentUser.user} (${roleLabel})`;
    userLabel.classList.remove("hidden");
    logoutBtn.classList.remove("hidden");
  } else if (userLabel && logoutBtn) {
    userLabel.classList.add("hidden");
    logoutBtn.classList.add("hidden");
  }

  if (page === "home") {
    const el = document.getElementById("page-home");
    if (el) {
      el.classList.remove("hidden");
      el.classList.add("active");
    }
  } else if (page === "admin") {
    const el = document.getElementById("page-admin");
    if (el) {
      el.classList.remove("hidden");
      el.classList.add("active");
    }
    Dashboard.renderAdmin();
  } else if (page === "developer") {
    Dashboard.renderDevHome();
  } else {
    Dashboard.renderBranchPage(page);
  }
}

document.getElementById("logoutBtn")?.addEventListener("click", async () => {
  if (
    await confirmModal(
      "تسجيل الخروج",
      "هل تريد تسجيل الخروج والعودة إلى القائمة الرئيسية؟",
    )
  ) {
    App.currentUser = null;
    App.currentBranch = null;
    sessionStorage.removeItem("currentSession");
    goToPage("home");
    showToast("تم تسجيل الخروج", "success");
  }
});
