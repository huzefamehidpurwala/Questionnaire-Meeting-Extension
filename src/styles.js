import { makeStyles, tokens } from "@fluentui/react-components";

export const useGeneralCss = makeStyles({
  gridCenterBox: {
    display: "grid",
    alignContent: "center",
    justifyContent: "center",
    // placeContent: "center",
  },

  flexCenterBox: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    // placeContent: "center",
  },

  border_0: {
    border: "0px",
  },
});

export const useAnalysisCss = makeStyles({
  clearBtn: {
    backgroundColor: "rgba(90,128,190, 1)",
    "&:hover": {
      backgroundColor: "rgba(90,128,190, 0.8)",
    },
    "&:hover:active": {
      backgroundColor: "rgba(90,128,190, 0.8)",
    },
    // width: "70%",
  },

  // checkBox: {},
});

export const useCreateQuestionnaireCss = makeStyles({
  questionnaireName: {
    // width: "75%",
    fontWeight: "bold",
  },
});

export const useQuestionnaireCss = makeStyles({
  startBtn: {
    backgroundColor: "rgba(16, 214, 123, 1)",
    "&:hover": {
      backgroundColor: "rgba(16, 214, 123, 0.8)",
    },
    "&:hover:active": {
      backgroundColor: "rgba(16, 214, 123, 1)",
    },
  },

  counter: {
    borderRadius: "1000px",
    backgroundColor: "rgba(84, 86, 118, 1)",
    height: "48px",
    width: "48px",
  },

  exitBtn: {
    backgroundColor: "rgba(235, 100, 100, 1)",
    "&:hover": {
      backgroundColor: "rgba(235, 100, 100, 0.8)",
    },
    "&:hover:active": {
      backgroundColor: "rgba(235, 100, 100, 1)",
    },
  },

  prevBtn: {
    backgroundColor: "rgba(235, 112, 159, 1)",
    "&:hover": {
      backgroundColor: "rgba(235, 112, 159, 0.8)",
    },
    "&:hover:active": {
      backgroundColor: "rgba(235, 112, 159, 1)",
    },
  },

  nextBtn: {
    backgroundColor: "rgba(156, 124, 242, 1)",
    "&:hover": {
      backgroundColor: "rgba(156, 124, 242, 0.8)",
    },
    "&:hover:active": {
      backgroundColor: "rgba(156, 124, 242, 1)",
    },
  },

  correctAns: {
    backgroundColor: tokens.colorStatusSuccessBackground3,
    color: "#fff",
  },

  wrongAns: {
    backgroundColor: tokens.colorStatusDangerBackground3,
    color: "#fff",
  },
});
