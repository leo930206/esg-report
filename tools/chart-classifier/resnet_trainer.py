"""
resnet_trainer.py  —  ESG 圖表分類 ResNet-50 訓練腳本 (v1.0)

工作流程：
  1. 先用 clip_labeler.py 產出 data/charts/{bar,line,pie,map,non_chart}/
  2. 手動清理錯誤標籤
  3. 執行本腳本 fine-tune ResNet-50，存最佳模型

使用方式：
  python resnet_trainer.py [--data_dir DATA_DIR] [--epochs 20] [--batch 32] [--lr 1e-4]

輸出：
  - models/resnet50_chart_best.pth   （最佳 val accuracy 的權重）
  - models/training_log.csv          （每 epoch 的 loss / acc）

依賴：
  pip install torch torchvision tqdm
"""

from __future__ import annotations

import argparse
import csv
import time
from pathlib import Path

import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader
from torchvision import datasets, models, transforms
from tqdm import tqdm


# ── 分類設定 ─────────────────────────────────────────────────────────────────

CATEGORIES = ["bar", "line", "pie", "map", "non_chart"]
NUM_CLASSES = len(CATEGORIES)


# ── 資料轉換 ─────────────────────────────────────────────────────────────────

TRAIN_TRANSFORMS = transforms.Compose([
    transforms.Resize((256, 256)),
    transforms.RandomCrop(224),
    transforms.RandomHorizontalFlip(),
    transforms.ColorJitter(brightness=0.2, contrast=0.2, saturation=0.1),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406],
                         std=[0.229, 0.224, 0.225]),
])

VAL_TRANSFORMS = transforms.Compose([
    transforms.Resize((224, 224)),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406],
                         std=[0.229, 0.224, 0.225]),
])


# ── 模型建立 ─────────────────────────────────────────────────────────────────

def build_model(num_classes: int, freeze_backbone: bool = False) -> nn.Module:
    model = models.resnet50(weights=models.ResNet50_Weights.IMAGENET1K_V2)
    if freeze_backbone:
        for name, param in model.named_parameters():
            if not name.startswith("layer4") and not name.startswith("fc"):
                param.requires_grad = False
    in_features = model.fc.in_features
    model.fc = nn.Linear(in_features, num_classes)
    return model


# ── 訓練主函式 ────────────────────────────────────────────────────────────────

def train(
    data_dir: Path,
    model_dir: Path,
    epochs: int,
    batch_size: int,
    lr: float,
    val_split: float,
    device: str,
    freeze_backbone: bool,
):
    # ── 資料集 ──
    # ImageFolder 讀 data/charts/{category}/ 結構
    full_dataset = datasets.ImageFolder(root=str(data_dir), transform=TRAIN_TRANSFORMS)
    class_names = full_dataset.classes
    print(f"偵測到 {len(class_names)} 個類別：{class_names}")
    print(f"總圖片數：{len(full_dataset)}")

    # train / val 分割（依 val_split 比例）
    n_val = max(1, int(len(full_dataset) * val_split))
    n_train = len(full_dataset) - n_val
    train_set, val_set = torch.utils.data.random_split(
        full_dataset, [n_train, n_val],
        generator=torch.Generator().manual_seed(42),
    )
    # val set 套用 val transform（建立新 dataset wrapper）
    val_set.dataset = datasets.ImageFolder(root=str(data_dir), transform=VAL_TRANSFORMS)

    train_loader = DataLoader(train_set, batch_size=batch_size, shuffle=True,
                              num_workers=4, pin_memory=True)
    val_loader   = DataLoader(val_set,   batch_size=batch_size, shuffle=False,
                              num_workers=4, pin_memory=True)

    print(f"訓練集：{n_train} 張，驗證集：{n_val} 張")

    # ── 模型 ──
    model = build_model(NUM_CLASSES, freeze_backbone=freeze_backbone).to(device)

    criterion = nn.CrossEntropyLoss()
    optimizer = optim.AdamW(
        filter(lambda p: p.requires_grad, model.parameters()),
        lr=lr, weight_decay=1e-4,
    )
    scheduler = optim.lr_scheduler.CosineAnnealingLR(optimizer, T_max=epochs)

    # ── 輸出目錄 ──
    model_dir.mkdir(parents=True, exist_ok=True)
    best_path = model_dir / "resnet50_chart_best.pth"
    log_path  = model_dir / "training_log.csv"

    best_val_acc = 0.0
    log_rows: list[dict] = []

    for epoch in range(1, epochs + 1):
        t0 = time.time()

        # ── 訓練 ──
        model.train()
        train_loss, train_correct, train_total = 0.0, 0, 0
        for images, labels in tqdm(train_loader, desc=f"Epoch {epoch}/{epochs} [train]", leave=False):
            images, labels = images.to(device), labels.to(device)
            optimizer.zero_grad()
            outputs = model(images)
            loss = criterion(outputs, labels)
            loss.backward()
            optimizer.step()

            train_loss    += loss.item() * images.size(0)
            preds          = outputs.argmax(dim=1)
            train_correct += (preds == labels).sum().item()
            train_total   += images.size(0)

        scheduler.step()

        # ── 驗證 ──
        model.eval()
        val_loss, val_correct, val_total = 0.0, 0, 0
        with torch.no_grad():
            for images, labels in tqdm(val_loader, desc=f"Epoch {epoch}/{epochs} [val]", leave=False):
                images, labels = images.to(device), labels.to(device)
                outputs = model(images)
                loss = criterion(outputs, labels)
                val_loss    += loss.item() * images.size(0)
                preds        = outputs.argmax(dim=1)
                val_correct += (preds == labels).sum().item()
                val_total   += images.size(0)

        t_epoch = time.time() - t0
        tr_loss = train_loss / train_total
        tr_acc  = train_correct / train_total * 100
        va_loss = val_loss / val_total
        va_acc  = val_correct / val_total * 100
        lr_now  = scheduler.get_last_lr()[0]

        print(
            f"Epoch {epoch:>3}/{epochs}  "
            f"train loss={tr_loss:.4f} acc={tr_acc:.1f}%  "
            f"val loss={va_loss:.4f} acc={va_acc:.1f}%  "
            f"lr={lr_now:.2e}  ({t_epoch:.0f}s)"
        )

        log_rows.append({
            "epoch": epoch,
            "train_loss": f"{tr_loss:.6f}",
            "train_acc":  f"{tr_acc:.2f}",
            "val_loss":   f"{va_loss:.6f}",
            "val_acc":    f"{va_acc:.2f}",
            "lr":         f"{lr_now:.8f}",
        })

        # ── 存最佳 ──
        if va_acc > best_val_acc:
            best_val_acc = va_acc
            torch.save({
                "epoch":      epoch,
                "model_state_dict": model.state_dict(),
                "val_acc":    va_acc,
                "class_names": class_names,
            }, best_path)
            print(f"  ✓ 已存最佳模型（val_acc={va_acc:.1f}%）→ {best_path}")

    # ── 寫 CSV log ──
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["epoch","train_loss","train_acc","val_loss","val_acc","lr"])
        writer.writeheader()
        writer.writerows(log_rows)

    print(f"\n訓練完成。最佳 val accuracy：{best_val_acc:.1f}%")
    print(f"模型：{best_path}")
    print(f"訓練記錄：{log_path}")


# ── 入口 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="ResNet-50 fine-tuning for ESG chart classification")
    parser.add_argument("--data_dir", type=str, default=None,
                        help="包含 bar/line/pie/map/non_chart 子目錄的路徑（預設：repo/data/charts/）")
    parser.add_argument("--model_dir", type=str, default=None,
                        help="模型輸出目錄（預設：repo/models/）")
    parser.add_argument("--epochs",   type=int,   default=20)
    parser.add_argument("--batch",    type=int,   default=32)
    parser.add_argument("--lr",       type=float, default=1e-4)
    parser.add_argument("--val_split", type=float, default=0.15,
                        help="驗證集比例（預設 0.15）")
    parser.add_argument("--freeze_backbone", action="store_true",
                        help="凍結 layer1-layer3，只訓練 layer4 + fc（資料少時使用）")
    args = parser.parse_args()

    # 自動偵測路徑
    repo_root = Path(__file__).resolve().parent.parent.parent
    data_dir  = Path(args.data_dir)  if args.data_dir  else repo_root / "data" / "charts"
    model_dir = Path(args.model_dir) if args.model_dir else repo_root / "models"

    # 裝置偵測
    if torch.cuda.is_available():
        device = "cuda"
    elif hasattr(torch.backends, "mps") and torch.backends.mps.is_available():
        device = "mps"
    else:
        device = "cpu"

    print(f"使用裝置：{device}")
    print(f"data_dir ：{data_dir}")
    print(f"model_dir：{model_dir}")

    # 驗證目錄結構
    if not data_dir.exists():
        print(f"錯誤：找不到 {data_dir}")
        print("請先執行 clip_labeler.py 產出分類結果，並完成手動清理後再執行本腳本。")
        return

    missing = [c for c in CATEGORIES if not (data_dir / c).is_dir()]
    if missing:
        print(f"警告：缺少類別目錄：{missing}")

    train(
        data_dir=data_dir,
        model_dir=model_dir,
        epochs=args.epochs,
        batch_size=args.batch,
        lr=args.lr,
        val_split=args.val_split,
        device=device,
        freeze_backbone=args.freeze_backbone,
    )


if __name__ == "__main__":
    main()
